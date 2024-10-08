﻿using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace Infinium.Modules.Marketing.Orders
{

    public class OrdersCalculate
    {
        private Infinium.Modules.Marketing.Orders.OrdersManager OrdersManager = null;

        public FrontsCalculate FrontsCalculate = null;
        public DecorCalculate DecorCalculate = null;

        public OrdersCalculate()
        {
            FrontsCalculate = new FrontsCalculate();
            DecorCalculate = new DecorCalculate(FrontsCalculate);
        }

        public OrdersCalculate(Infinium.Modules.Marketing.Orders.OrdersManager tOrdersManager)
        {
            OrdersManager = tOrdersManager;

            FrontsCalculate = new FrontsCalculate();
            DecorCalculate = new DecorCalculate(FrontsCalculate);
        }

        public void GetMegaOrderDiscount(int MeganOrderID, ref int CurrencyTypeID, ref decimal PaymentRate,
            ref decimal ProfilDiscountDirector, ref decimal TPSDiscountDirector, ref decimal ProfilTotalDiscount, ref decimal TPSTotalDiscount)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ProfilDiscountDirector, TPSDiscountDirector, ProfilTotalDiscount, TPSTotalDiscount, CurrencyTypeID, PaymentRate FROM MegaOrders WHERE MeganOrderID = " + MeganOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows.Count > 0)
                    {
                        ProfilDiscountDirector = Convert.ToDecimal(DT.Rows[0]["ProfilDiscountDirector"]);
                        TPSDiscountDirector = Convert.ToDecimal(DT.Rows[0]["TPSDiscountDirector"]);
                        ProfilTotalDiscount = Convert.ToDecimal(DT.Rows[0]["ProfilTotalDiscount"]) - Convert.ToDecimal(DT.Rows[0]["ProfilDiscountDirector"]);
                        TPSTotalDiscount = Convert.ToDecimal(DT.Rows[0]["TPSTotalDiscount"]) - Convert.ToDecimal(DT.Rows[0]["TPSDiscountDirector"]);
                        CurrencyTypeID = Convert.ToInt32(DT.Rows[0]["CurrencyTypeID"]);
                        PaymentRate = Convert.ToDecimal(DT.Rows[0]["PaymentRate"]);
                    }
                }
            }
        }

        public void GetMainOrderTotalSum(int MegaOrderID, ref decimal ProfilTotalSum, ref decimal TPSTotalSum)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MainOrderID, OriginalProfilOrderCost, OriginalTPSOrderCost FROM MainOrders WHERE MegaOrderID = " + MegaOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        ProfilTotalSum += Convert.ToDecimal(DT.Rows[i]["OriginalProfilOrderCost"]);
                        TPSTotalSum += Convert.ToDecimal(DT.Rows[i]["OriginalTPSOrderCost"]);
                    }
                }
            }
        }

        public int GetClientID(int MainOrderID)
        {
            int ClientID = 0;

            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientID FROM MegaOrders" +
                    " WHERE MegaOrderID = (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID = " + MainOrderID + ")",
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        ClientID = Convert.ToInt32(DT.Rows[0]["ClientID"]);
                }
            }

            return ClientID;
        }

        public void CalculateOrder(int MegaOrderID, int MainOrderID,
            decimal ProfilDiscountDirector, decimal TPSDiscountDirector, decimal ProfilTotalDiscount, decimal TPSTotalDiscount, decimal DiscountPaymentCondition, int CurrencyTypeID, decimal Currency,
            object ConfirmDateTime,
            bool NeedCalcPrice)
        {
            int ClientID = GetClientID(MainOrderID);
            DecorCalculate.GetAllDecorOrders(MegaOrderID);
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DocDateTime, MainOrderID, ProfilTotalDiscount, TPSTotalDiscount, FrontsSquare, FrontsCost," +
                " DecorCost, ProfilOrderCost, TPSOrderCost, OrderCost, OriginalProfilOrderCost, OriginalTPSOrderCost, OriginalOrderCost, CurrencyOrderCost, CurrencyTypeID, Weight, TechDrilling FROM MainOrders WHERE MainOrderID = " + MainOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);

                        decimal FrontsOrderSquare = 0;

                        decimal OriginalProfilFrontsOrderCost = 0;
                        decimal OriginalTPSFrontsOrderCost = 0;
                        decimal OriginalTotalFrontsOrderCost = 0;
                        decimal ProfilFrontsOrderCost = 0;
                        decimal TPSFrontsOrderCost = 0;
                        decimal TotalFrontsOrderCost = 0;

                        decimal OriginalProfilDecorOrderCost = 0;
                        decimal OriginalTPSDecorOrderCost = 0;
                        decimal OriginalTotalDecorOrderCost = 0;
                        decimal ProfilDecorOrderCost = 0;
                        decimal TPSDecorOrderCost = 0;
                        decimal TotalDecorOrderCost = 0;

                        decimal OriginalProfilOrderCost = 0;
                        decimal OriginalTPSOrderCost = 0;
                        decimal OriginalOrderCost = 0;
                        decimal ProfilOrderCost = 0;
                        decimal TPSOrderCost = 0;
                        decimal OrderCost = 0;
                        decimal CurrencyOrderCost = 0;

                        decimal FrontsOrderWeight = 0;
                        decimal DecorOrderWeight = 0;
                        decimal TotalWeight = 0;
                        Currency = Decimal.Round(Currency, 4, MidpointRounding.AwayFromZero);
                        FrontsCalculate.NeedCalcPrice = NeedCalcPrice;
                        bool TechDrilling = Convert.ToBoolean(DT.Rows[0]["TechDrilling"]);
                        FrontsCalculate.CalculateFronts(ClientID, MainOrderID, ProfilDiscountDirector, TPSDiscountDirector, ProfilTotalDiscount, TPSTotalDiscount, CurrencyTypeID, Currency, TechDrilling,
                            ref OriginalProfilFrontsOrderCost, ref OriginalTPSFrontsOrderCost, ref OriginalTotalFrontsOrderCost,
                            ref ProfilFrontsOrderCost, ref TPSFrontsOrderCost, ref TotalFrontsOrderCost, ref CurrencyOrderCost,
                            ref FrontsOrderSquare, ref FrontsOrderWeight, ConfirmDateTime);
                        DecorCalculate.NeedCalcPrice = NeedCalcPrice;
                        DecorCalculate.CalculateDecor(ClientID, MainOrderID, ProfilDiscountDirector, TPSDiscountDirector, ProfilTotalDiscount, TPSTotalDiscount, CurrencyTypeID, Currency,
                            ref OriginalProfilDecorOrderCost, ref OriginalTPSDecorOrderCost, ref OriginalTotalDecorOrderCost,
                            ref ProfilDecorOrderCost, ref TPSDecorOrderCost, ref TotalDecorOrderCost, ref CurrencyOrderCost,
                            ref DecorOrderWeight, ConfirmDateTime);

                        ProfilFrontsOrderCost = Decimal.Round(ProfilFrontsOrderCost, 2, MidpointRounding.AwayFromZero);
                        TPSFrontsOrderCost = Decimal.Round(TPSFrontsOrderCost, 2, MidpointRounding.AwayFromZero);
                        TotalFrontsOrderCost = Decimal.Round(TotalFrontsOrderCost, 2, MidpointRounding.AwayFromZero);

                        ProfilDecorOrderCost = Decimal.Round(ProfilDecorOrderCost, 2, MidpointRounding.AwayFromZero);
                        TPSDecorOrderCost = Decimal.Round(TPSDecorOrderCost, 2, MidpointRounding.AwayFromZero);
                        TotalDecorOrderCost = Decimal.Round(TotalDecorOrderCost, 2, MidpointRounding.AwayFromZero);

                        OriginalProfilOrderCost = OriginalProfilFrontsOrderCost + OriginalProfilDecorOrderCost;
                        OriginalTPSOrderCost = OriginalTPSFrontsOrderCost + OriginalTPSDecorOrderCost;
                        OriginalOrderCost = OriginalProfilOrderCost + OriginalTPSOrderCost;

                        ProfilOrderCost = ProfilFrontsOrderCost + ProfilDecorOrderCost;
                        TPSOrderCost = TPSFrontsOrderCost + TPSDecorOrderCost;
                        OrderCost = ProfilOrderCost + TPSOrderCost;

                        decimal PackWeight = 0;

                        if (FrontsOrderSquare > 0)
                            PackWeight = FrontsOrderSquare * Convert.ToDecimal(0.7);

                        TotalWeight = Decimal.Round(FrontsOrderWeight + PackWeight + DecorOrderWeight, 3, MidpointRounding.AwayFromZero);

                        decimal ProfilDiscountOrderSum = ProfilTotalDiscount - DiscountPaymentCondition;
                        decimal TPSDiscountOrderSum = TPSTotalDiscount - DiscountPaymentCondition;
                        decimal dP = 1 - (1 - DiscountPaymentCondition / 100) * (1 - ProfilDiscountOrderSum / 100) * (1 - ProfilDiscountDirector / 100);
                        decimal dT = 1 - (1 - DiscountPaymentCondition / 100) * (1 - TPSDiscountOrderSum / 100) * (1 - TPSDiscountDirector / 100);

                        dP = Decimal.Round(dP * 100, 2, MidpointRounding.AwayFromZero);
                        dT = Decimal.Round(dT * 100, 2, MidpointRounding.AwayFromZero);

                        DT.Rows[0]["ProfilTotalDiscount"] = dP;
                        DT.Rows[0]["TPSTotalDiscount"] = dT;
                        DT.Rows[0]["FrontsSquare"] = Decimal.Round(FrontsOrderSquare, 3, MidpointRounding.AwayFromZero);
                        DT.Rows[0]["FrontsCost"] = TotalFrontsOrderCost;
                        DT.Rows[0]["DecorCost"] = TotalDecorOrderCost;
                        DT.Rows[0]["OriginalProfilOrderCost"] = OriginalProfilOrderCost;
                        DT.Rows[0]["OriginalTPSOrderCost"] = OriginalTPSOrderCost;
                        DT.Rows[0]["OriginalOrderCost"] = OriginalOrderCost;
                        DT.Rows[0]["ProfilOrderCost"] = ProfilOrderCost;
                        DT.Rows[0]["TPSOrderCost"] = TPSOrderCost;
                        DT.Rows[0]["OrderCost"] = OrderCost;
                        DT.Rows[0]["CurrencyTypeID"] = CurrencyTypeID;
                        DT.Rows[0]["CurrencyOrderCost"] = CurrencyOrderCost;
                        DT.Rows[0]["Weight"] = TotalWeight;

                        DA.Update(DT);
                    }
                }
            }
        }

        public void Recalculate(int MegaOrderID,
            decimal ProfilDiscountDirector, decimal TPSDiscountDirector, decimal ProfilTotalDiscount, decimal TPSTotalDiscount, decimal DiscountPaymentCondition,
            int CurrencyTypeID, decimal PaymentRate, object ConfirmDateTime, decimal CurrencyTotalCost, bool NeedCalcPrice)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MainOrderID FROM MainOrders WHERE MegaOrderID = " + MegaOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        int MainOrderID = Convert.ToInt32(DT.Rows[i]["MainOrderID"]);
                        CalculateOrder(MegaOrderID, MainOrderID, ProfilDiscountDirector, TPSDiscountDirector, ProfilTotalDiscount, TPSTotalDiscount, DiscountPaymentCondition, CurrencyTypeID, PaymentRate, ConfirmDateTime, NeedCalcPrice);
                    }
                }
            }

            SummaryCost(MegaOrderID, CurrencyTotalCost);
        }

        public void SummaryCost(int MegaOrderID,
            decimal CurrencyTotalCost)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM MainOrders WHERE MegaOrderID = " + MegaOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    decimal OrderCost = 0;
                    decimal TotalWeight = 0;
                    decimal TotalSquare = 0;
                    decimal TotalCost = 0;
                    decimal ProfilTotalDiscount = 0;
                    decimal TPSTotalDiscount = 0;
                    //decimal CurrencyOrderCost = 0;

                    int Factory = -1;

                    foreach (DataRow Row in DT.Rows)
                    {
                        ProfilTotalDiscount = Convert.ToDecimal(Row["ProfilTotalDiscount"]);
                        TPSTotalDiscount = Convert.ToDecimal(Row["TPSTotalDiscount"]);
                        OrderCost += Convert.ToDecimal(Row["OrderCost"]);
                        //CurrencyOrderCost += Convert.ToDecimal(Row["CurrencyOrderCost"]);
                        TotalWeight += Convert.ToDecimal(Row["Weight"]);
                        TotalSquare += Convert.ToDecimal(Row["FrontsSquare"]);

                        if (Factory == 0)
                            continue;

                        if (Row["FactoryID"].ToString() == "0")
                            Factory = 0;

                        if (Row["FactoryID"].ToString() == "1")
                            if (Factory == 2)
                                Factory = 0;
                            else
                                Factory = 1;

                        if (Row["FactoryID"].ToString() == "2")
                            if (Factory == 1)
                                Factory = 0;
                            else
                                Factory = 2;
                    }

                    using (SqlDataAdapter MegaDA = new SqlDataAdapter("SELECT MegaOrderID, OrderCost, CurrencyTransportCost, CurrencyAdditionalCost, CurrencyOrderCost, CurrencyTotalCost, Weight, Square," +
                        " FactoryID, TotalCost, ProfilTotalDiscount, TPSTotalDiscount, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost" +
                        " FROM MegaOrders WHERE MegaOrderID = " + MegaOrderID,
                        ConnectionStrings.MarketingOrdersConnectionString))
                    {
                        using (SqlCommandBuilder CB = new SqlCommandBuilder(MegaDA))
                        {
                            using (DataTable MegaDT = new DataTable())
                            {
                                MegaDA.Fill(MegaDT);
                                decimal ComplaintProfilCost = Convert.ToDecimal(MegaDT.Rows[0]["ComplaintProfilCost"]);
                                decimal ComplaintTPSCost = Convert.ToDecimal(MegaDT.Rows[0]["ComplaintTPSCost"]);
                                decimal TransportCost = Convert.ToDecimal(MegaDT.Rows[0]["TransportCost"]);
                                decimal AdditionalCost = Convert.ToDecimal(MegaDT.Rows[0]["AdditionalCost"]);
                                decimal TotalProfilAdditionalCost = -ComplaintProfilCost + TransportCost + AdditionalCost;
                                decimal TotalTPSAdditionalCost = -ComplaintTPSCost + TransportCost + AdditionalCost;
                                decimal CurrencyTransportCost = Convert.ToDecimal(MegaDT.Rows[0]["CurrencyTransportCost"]);
                                decimal CurrencyAdditionalCost = Convert.ToDecimal(MegaDT.Rows[0]["CurrencyAdditionalCost"]);

                                decimal CurrencyOrderCost = Decimal.Round((CurrencyTotalCost - CurrencyTransportCost - CurrencyAdditionalCost), 2, MidpointRounding.AwayFromZero);

                                MegaDT.Rows[0]["CurrencyTotalCost"] = CurrencyTotalCost;
                                MegaDT.Rows[0]["CurrencyOrderCost"] = CurrencyOrderCost;

                                FrontsCalculate.AssignTransportCost(MegaOrderID, -ComplaintProfilCost, -ComplaintTPSCost, (TransportCost + AdditionalCost));
                                DecorCalculate.AssignTransportCost(MegaOrderID, -ComplaintProfilCost, -ComplaintTPSCost, (TransportCost + AdditionalCost));

                                TotalWeight = Decimal.Round(TotalWeight, 3, MidpointRounding.AwayFromZero);
                                TotalSquare = Decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);

                                //MegaDT.Rows[0]["ProfilTotalDiscount"] = ProfilTotalDiscount;
                                //MegaDT.Rows[0]["TPSTotalDiscount"] = TPSTotalDiscount;
                                MegaDT.Rows[0]["OrderCost"] = OrderCost;
                                //MegaDT.Rows[0]["CurrencyOrderCost"] = CurrencyOrderCost;
                                MegaDT.Rows[0]["Weight"] = TotalWeight;
                                MegaDT.Rows[0]["Square"] = TotalSquare;
                                MegaDT.Rows[0]["FactoryID"] = Factory;

                                TotalCost = OrderCost + TransportCost + AdditionalCost;
                                MegaDT.Rows[0]["TotalCost"] = TotalCost;

                                MegaDA.Update(MegaDT);
                            }
                        }
                    }
                }
            }
        }

        public void SetProfilCurrencyCost(int MegaOrderID, decimal ProfilTotalDiscount, decimal TPSTotalDiscount, decimal ProfilDiscountOrderSum, decimal TPSDiscountOrderSum, decimal PaymentRate)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrderID, ProfilTotalDiscount,TPSTotalDiscount, ProfilDiscountOrderSum, TPSDiscountOrderSum, PaymentRate, FactoryID" +
                " FROM MegaOrders WHERE MegaOrderID = " + MegaOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);
                        if (DT.Rows.Count > 0)
                        {
                            DT.Rows[0]["ProfilTotalDiscount"] = ProfilTotalDiscount;
                            DT.Rows[0]["ProfilDiscountOrderSum"] = ProfilDiscountOrderSum;
                            //DT.Rows[0]["PaymentRate"] = PaymentRate;
                        }
                        DA.Update(DT);
                    }
                }
            }
        }

        public void SetTPSCurrencyCost(int MegaOrderID, decimal ProfilTotalDiscount, decimal TPSTotalDiscount, decimal ProfilDiscountOrderSum, decimal TPSDiscountOrderSum, decimal PaymentRate)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrderID, ProfilTotalDiscount,TPSTotalDiscount, ProfilDiscountOrderSum, TPSDiscountOrderSum, PaymentRate, FactoryID" +
                " FROM MegaOrders WHERE MegaOrderID = " + MegaOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);
                        if (DT.Rows.Count > 0)
                        {
                            DT.Rows[0]["TPSTotalDiscount"] = TPSTotalDiscount;
                            DT.Rows[0]["TPSDiscountOrderSum"] = TPSDiscountOrderSum;
                            //DT.Rows[0]["PaymentRate"] = PaymentRate;
                        }
                        DA.Update(DT);
                    }
                }
            }
        }

        public void GetDecorExcluziveForManager(int ManagerID)
        {
            DataTable DecorConfigDataTable = new DataTable();
            DataTable ExcluziveCatalogDataTable = new DataTable();
            DataTable ClientsDataTable = new DataTable();

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT dbo.DecorConfig.DecorConfigID, dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS DecorName, Colors.TechStoreName AS ColorName, dbo.Patina.PatinaName AS PatinaName FROM DecorConfig INNER JOIN
                dbo.DecorProducts ON dbo.DecorConfig.ProductID = dbo.DecorProducts.ProductID INNER JOIN
                dbo.TechStore AS Decor ON dbo.DecorConfig.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                dbo.TechStore AS Colors ON dbo.DecorConfig.ColorID = Colors.TechStoreID INNER JOIN
                dbo.Patina ON dbo.DecorConfig.PatinaID = dbo.Patina.PatinaID
                WHERE DecorConfigID IN (SELECT ConfigID FROM infiniu2_marketingreference.dbo.ExcluziveCatalog
                WHERE ClientID IN (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients WHERE ManagerID = " + ManagerID + " AND ProductType = 1))", ConnectionStrings.CatalogConnectionString))
                DA.Fill(DecorConfigDataTable);
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM ExcluziveCatalog 
                WHERE ClientID IN (SELECT ClientID FROM dbo.Clients WHERE ManagerID = " + ManagerID + ") AND ProductType = 1", ConnectionStrings.MarketingReferenceConnectionString))
                DA.Fill(ExcluziveCatalogDataTable);
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT ClientID FROM Clients WHERE ManagerID = " + ManagerID + " ORDER BY ClientName", ConnectionStrings.MarketingReferenceConnectionString))
                DA.Fill(ClientsDataTable);

            DataTable ResultDataTable = new DataTable();
            ResultDataTable.Columns.Add(new DataColumn("DecorConfigID", Type.GetType("System.Int32")));
            ResultDataTable.Columns.Add(new DataColumn("ProductName", Type.GetType("System.String")));
            ResultDataTable.Columns.Add(new DataColumn("DecorName", Type.GetType("System.String")));
            ResultDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            ResultDataTable.Columns.Add(new DataColumn("PatinaName", Type.GetType("System.String")));
            ResultDataTable.Columns.Add(new DataColumn("Price", Type.GetType("System.Decimal")));
            for (int i = 0; i < ClientsDataTable.Rows.Count; i++)
            {
                Int64 ClientID = Convert.ToInt64(ClientsDataTable.Rows[i]["ClientID"]);
                ResultDataTable.Columns.Add(new DataColumn("ClientID" + ClientID, Type.GetType("System.Decimal")));
            }
            for (int i = 0; i < DecorConfigDataTable.Rows.Count; i++)
            {
                DataRow NewRow = ResultDataTable.NewRow();
                NewRow["DecorConfigID"] = Convert.ToInt32(DecorConfigDataTable.Rows[i]["DecorConfigID"]);
                NewRow["ProductName"] = DecorConfigDataTable.Rows[i]["ProductName"].ToString();
                NewRow["DecorName"] = DecorConfigDataTable.Rows[i]["DecorName"].ToString();
                NewRow["ColorName"] = DecorConfigDataTable.Rows[i]["ColorName"].ToString();
                NewRow["PatinaName"] = DecorConfigDataTable.Rows[i]["PatinaName"].ToString();
                ResultDataTable.Rows.Add(NewRow);
            }
            for (int i = 0; i < ResultDataTable.Rows.Count; i++)
            {
                int DecorConfigID = Convert.ToInt32(ResultDataTable.Rows[i]["DecorConfigID"]);
                for (int j = 0; j < ClientsDataTable.Rows.Count; j++)
                {
                    int ClientID = Convert.ToInt32(ClientsDataTable.Rows[j]["ClientID"]);
                    DataRow[] rows = ExcluziveCatalogDataTable.Select("ConfigID=" + DecorConfigID + " AND ClientID=" + ClientID);
                    if (rows.Count() > 0)
                        ResultDataTable.Rows[i]["ClientID" + ClientID] = ResultDataTable.Rows[i]["Price"];
                }
            }
        }

    }










    public class FrontsCalculate
    {
        private bool Sale = false;
        public bool NeedCalcPrice = false;

        private DataTable StockDetailsDT = null;
        private DataTable FrontsConfigDataTable = null;
        private DataTable DecorConfigDataTable = null;
        private DataTable TechStoreDataTable = null;
        private DataTable StandardDataTable = null;
        private DataTable MeasuresDataTable = null;
        private DataTable InsetTypesDataTable = null;
        private DataTable InsetPriceDataTable = null;
        private DataTable ClientsDataTable = null;
        private DataTable FrontsDataTable = null;
        private DataTable AluminiumFrontsDataTable = null;
        private DataTable FrontsDiscountTypesDataTable = null;
        private DataTable InsetMarginsDataTable = null;
        private DataTable FrontsDiscountDataTable = null;

        public FrontsCalculate()
        {
            Initialize();
        }


        private void Create()
        {

        }

        private void Fill()
        {
            string SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore 
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
            FrontsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }

            InsetTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }

            AluminiumFrontsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM AluminiumFronts",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(AluminiumFrontsDataTable);
            }

            StandardDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Standard",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(StandardDataTable);
            }

            FrontsConfigDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsConfig",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(FrontsConfigDataTable);
            //}
            FrontsConfigDataTable = TablesManager.FrontsConfigDataTableAll;

            DecorConfigDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorConfig",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDataTable);
            //}
            DecorConfigDataTable = TablesManager.DecorConfigDataTableAll;
            TechStoreDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM TechStore",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(TechStoreDataTable);
            //}
            TechStoreDataTable = TablesManager.TechStoreDataTable;

            MeasuresDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Measures",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(MeasuresDataTable);
            }

            InsetMarginsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetMargins",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetMarginsDataTable);
            }

            InsetPriceDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetPrice",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(InsetPriceDataTable);
            }

            FrontsDiscountDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DiscountFronts",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(FrontsDiscountDataTable);
            }

            FrontsDiscountTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsDiscountTypes",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(FrontsDiscountTypesDataTable);
            }

            ClientsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Clients",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsDataTable);
            }
        }

        private void Initialize()
        {
            Create();
            Fill();
        }

        private void ProfilFrontClientPrice(int ClientID, DataRow FrontsOrdersRow, decimal DiscountDirector, decimal TotalDiscount, bool TechDrilling,
            ref decimal OriginalPrice, ref decimal Price)
        {
            decimal MarketingCost = 0;//себестоимость
            decimal PriceRatio = 0;//ценовой коэффициент
            int DigitCapacity = 0;//разрядность
            decimal PriceGroup = 0;//ценовая группа

            DataRow[] FPRow = ClientsDataTable.Select("ClientID = " + ClientID);
            if (FPRow.Count() > 0)
            {
                bool IsSample = Convert.ToBoolean(FrontsOrdersRow["IsSample"]);//Образец
                PriceGroup = Convert.ToDecimal(FPRow[0]["PriceGroup"]);//MarketingCost
                DataRow[] PRows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());
                if (PRows.Count() > 0)
                {
                    MarketingCost = Convert.ToDecimal(PRows[0]["MarketingCost"]);
                    PriceRatio = Convert.ToDecimal(PRows[0]["PriceRatio"]);
                    DigitCapacity = Convert.ToInt32(PRows[0]["DigitCapacity"]);
                    OriginalPrice = MarketingCost * PriceRatio * (1 + PriceGroup);
                    OriginalPrice = Decimal.Round(OriginalPrice, DigitCapacity, MidpointRounding.AwayFromZero);

                    if (ClientID == 145)
                        OriginalPrice = Convert.ToDecimal(PRows[0]["ZOVRetailPrice"]);
                    if (Sale)
                    {
                        decimal SampleSaleValue = GetSampleDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                        decimal SaleValue = GetDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                        if (IsSample)
                            Price = (OriginalPrice - OriginalPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                        else
                            Price = (OriginalPrice - OriginalPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                    }
                    else
                    {
                        if (IsSample)
                            Price = (OriginalPrice - OriginalPrice / 100 * 50) * (1 - DiscountDirector / 100);
                        else
                            Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    }
                    //if (IsSample)
                    //    Price = (OriginalPrice - OriginalPrice * 50 / 100) * (1 - DiscountDirector / 100);
                    //else
                    //    Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    if (TechDrilling)
                        Price = Price * 1.3m;
                    Price = Decimal.Round(Price, 2, MidpointRounding.AwayFromZero);
                }
            }
        }

        private void TPSFrontClientPrice(int ClientID, DataRow FrontsOrdersRow, decimal DiscountDirector, decimal TotalDiscount, bool TechDrilling,
            ref decimal OriginalPrice, ref decimal Price)
        {
            decimal MarketingCost = 0;//себестоимость
            decimal PriceRatio = 0;//ценовой коэффициент
            int DigitCapacity = 0;//разрядность
            decimal PriceGroup = 0;//ценовая группа

            DataRow[] FPRow = ClientsDataTable.Select("ClientID = " + ClientID);
            if (FPRow.Count() > 0)
            {
                bool IsSample = Convert.ToBoolean(FrontsOrdersRow["IsSample"]);//Образец
                PriceGroup = Convert.ToDecimal(FPRow[0]["PriceGroup"]);//MarketingCost
                DataRow[] PRows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());
                if (PRows.Count() > 0)
                {
                    MarketingCost = Convert.ToDecimal(PRows[0]["MarketingCost"]);
                    PriceRatio = Convert.ToDecimal(PRows[0]["PriceRatio"]);
                    DigitCapacity = Convert.ToInt32(PRows[0]["DigitCapacity"]);
                    OriginalPrice = MarketingCost * PriceRatio * (1 + PriceGroup);
                    OriginalPrice = Decimal.Round(OriginalPrice, DigitCapacity, MidpointRounding.AwayFromZero);

                    if (ClientID == 145)
                        OriginalPrice = Convert.ToDecimal(PRows[0]["ZOVRetailPrice"]);
                    if (Sale)
                    {
                        decimal SampleSaleValue = GetSampleDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                        decimal SaleValue = GetDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                        if (IsSample)
                            Price = (OriginalPrice - OriginalPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                        else
                            Price = (OriginalPrice - OriginalPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                    }
                    else
                    {
                        if (IsSample)
                            Price = (OriginalPrice - OriginalPrice / 100 * 50) * (1 - DiscountDirector / 100);
                        else
                            Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    }
                    //if (IsSample)
                    //{
                    //    Price = (OriginalPrice - OriginalPrice * 50 / 100) * (1 - DiscountDirector / 100);
                    //}
                    //else
                    //{
                    //    Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    //}
                    if (TechDrilling)
                        Price = Price * 1.3m;
                    Price = Decimal.Round(Price, 2, MidpointRounding.AwayFromZero);
                }
            }
        }

        private bool IsStandardSize(DataRow FrontsOrdersRow)
        {
            DataRow[] Rows = StandardDataTable.Select("Height = " + FrontsOrdersRow["Height"].ToString() +
                                                      " AND Width = " + FrontsOrdersRow["Width"].ToString());

            //return (Rows.Count() > 0);
            return true;
        }

        private void ProfilFrontPrice(int ClientID, DataRow FrontsOrdersRow, decimal DiscountDirector, decimal TotalDiscount, 
            bool TechDrilling, ref decimal OriginalPrice, ref decimal Price, decimal Currency)
        {
            bool IsNonStandard = false;
            decimal ExtraPrice = 0;

            decimal p = 0;
            ProfilFrontClientPrice(ClientID, FrontsOrdersRow, DiscountDirector, TotalDiscount, TechDrilling, ref OriginalPrice, ref p);

            DataRow[] Row = ClientsDataTable.Select("ClientID = " + ClientID);
            if (Row.Count() > 0)
            {
                if (Convert.ToBoolean(FrontsOrdersRow["IsNonStandard"]) ||
                    Convert.ToBoolean(Row[0]["NonStandard"]) == true)
                    ExtraPrice = GetNonStandardMargin(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]), ClientID);
                else
                    ExtraPrice = 0;

                //если фасад нестандартный и имеет наценку и прямой
                if (Convert.ToBoolean(FrontsOrdersRow["IsNonStandard"]) ||
                    (ExtraPrice > 0 && IsStandardSize(FrontsOrdersRow) == false && Convert.ToBoolean(Row[0]["NonStandard"]) == true))
                {
                    Price = p / 100 * ExtraPrice + p;
                    IsNonStandard = true;
                }
                else
                {
                    Price = p;
                    IsNonStandard = false;
                }
            }
            FrontsOrdersRow["ClientOriginalPrice"] = Math.Ceiling(OriginalPrice * Currency / 0.01m) * 0.01m;
            FrontsOrdersRow["OriginalPrice"] = OriginalPrice;
            FrontsOrdersRow["IsNonStandard"] = IsNonStandard;
            FrontsOrdersRow["FrontPrice"] = Price;
            bool IsSample = Convert.ToBoolean(FrontsOrdersRow["IsSample"]);
            //if (IsSample)
            //{
            //    if (OriginalPrice != 0)
            //        FrontsOrdersRow["TotalDiscount"] = (1 - Price / OriginalPrice) * 100;
            //}
            //else
            //    FrontsOrdersRow["TotalDiscount"] = TotalDiscount + DiscountDirector;
            if (OriginalPrice != 0)
            {
                decimal d = Decimal.Round((1 - Price / OriginalPrice) * 100, 2, MidpointRounding.AwayFromZero);
                FrontsOrdersRow["TotalDiscount"] = d;
            }
        }

        private void TPSFrontPrice(int ClientID, DataRow FrontsOrdersRow, decimal DiscountDirector, decimal TotalDiscount, bool TechDrilling, ref decimal OriginalPrice, ref decimal Price)
        {
            bool IsNonStandard = false;
            decimal ExtraPrice = 0;
            int FrontConfigID = Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]);
            decimal p = 0;
            TPSFrontClientPrice(ClientID, FrontsOrdersRow, DiscountDirector, TotalDiscount, TechDrilling, ref OriginalPrice, ref p);

            DataRow[] Row = ClientsDataTable.Select("ClientID = " + ClientID);
            if (Row.Count() > 0)
            {
                if (Convert.ToBoolean(FrontsOrdersRow["IsNonStandard"]) ||
                    Convert.ToBoolean(Row[0]["NonStandard"]) == true)
                    ExtraPrice = GetNonStandardMargin(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]), ClientID);
                else
                    ExtraPrice = 0;

                //если фасад нестандартный и имеет наценку и прямой
                if (Convert.ToBoolean(FrontsOrdersRow["IsNonStandard"]) ||
                    (ExtraPrice > 0 && IsStandardSize(FrontsOrdersRow) == false && Convert.ToBoolean(Row[0]["NonStandard"]) == true))
                {
                    Price = p / 100 * ExtraPrice + p;
                    IsNonStandard = true;
                }
                else
                {
                    Price = p;
                    IsNonStandard = false;
                }

            }
            FrontsOrdersRow["IsNonStandard"] = IsNonStandard;
            FrontsOrdersRow["OriginalPrice"] = OriginalPrice;
            FrontsOrdersRow["FrontPrice"] = Price;
            bool IsSample = Convert.ToBoolean(FrontsOrdersRow["IsSample"]);
            //if (IsSample)
            //{
            //    if (OriginalPrice != 0)
            //        FrontsOrdersRow["TotalDiscount"] = (1 - Price / OriginalPrice) * 100;
            //}
            //else
            //    FrontsOrdersRow["TotalDiscount"] = TotalDiscount + DiscountDirector;
            if (OriginalPrice != 0)
            {
                decimal d = Decimal.Round((1 - Price / OriginalPrice) * 100, 2, MidpointRounding.AwayFromZero);
                FrontsOrdersRow["TotalDiscount"] = d;
            }
        }

        private void InsetPrice(int ClientID, DataRow FrontsOrdersRow, decimal DiscountDirector, decimal TotalDiscount, bool TechDrilling, ref decimal OriginalPrice, ref decimal Price)
        {
            decimal MarketingCost = 0;//себестоимость
            decimal PriceRatio = 0;//ценовой коэффициент
            int DigitCapacity = 0;//разрядность
            decimal PriceGroup = 0;//ценовая группа
            int FrontID = Convert.ToInt32(FrontsOrdersRow["FrontID"]);
            int InsetTypeID = Convert.ToInt32(FrontsOrdersRow["InsetTypeID"]);

            bool IsSample = Convert.ToBoolean(FrontsOrdersRow["IsSample"]);//Образец
            //Решетки
            if (InsetTypeID == 685 || InsetTypeID == 686 || InsetTypeID == 687 || InsetTypeID == 688 || InsetTypeID == 29470 || InsetTypeID == 29471)
            {
                DataRow[] Rows = DecorConfigDataTable.Select("DecorID = " + InsetTypeID);
                if (Rows.Count() > 0)
                {
                    DataRow[] CPRow = ClientsDataTable.Select("ClientID = " + ClientID);
                    if (CPRow.Count() > 0)
                    {
                        PriceGroup = Convert.ToDecimal(CPRow[0]["PriceGroup"]);//MarketingCost
                        MarketingCost = Convert.ToDecimal(Rows[0]["MarketingCost"]);
                        PriceRatio = Convert.ToDecimal(Rows[0]["PriceRatio"]);
                        DigitCapacity = Convert.ToInt32(Rows[0]["DigitCapacity"]);
                        OriginalPrice = MarketingCost * PriceRatio * (1 + PriceGroup);
                        OriginalPrice = Decimal.Round(OriginalPrice, DigitCapacity, MidpointRounding.AwayFromZero);

                        if (ClientID == 145)
                            OriginalPrice = Convert.ToDecimal(Rows[0]["ZOVPrice"]);
                        if (Sale)
                        {
                            decimal SampleSaleValue = GetSampleDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                            decimal SaleValue = GetDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                            if (IsSample)
                                Price = (OriginalPrice - OriginalPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                            else
                                Price = (OriginalPrice - OriginalPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                        }
                        else
                        {
                            if (IsSample)
                                Price = (OriginalPrice - OriginalPrice / 100 * 50) * (1 - DiscountDirector / 100);
                            else
                                Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                        }
                        if (TechDrilling)
                            Price = Price * 1.3m;
                        //if (IsSample)
                        //{
                        //    Price = (OriginalPrice - OriginalPrice * 50 / 100) * (1 - DiscountDirector / 100);
                        //    //Price = OriginalPrice - OriginalPrice / 100 * (50 + DiscountDirector);
                        //}
                        //else
                        //{
                        //    Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                        //    //Price = CurrentPrice - CurrentPrice / 100 * (TotalDiscount + DiscountDirector);
                        //}
                        FrontsOrdersRow["OriginalInsetPrice"] = OriginalPrice;
                        FrontsOrdersRow["InsetPrice"] = Price;
                        return;
                    }
                }
            }
            //Стекло
            if (IsAluminium(FrontsOrdersRow) > -1 && InsetTypeID == 2)
            {
                int InsetColorID = Convert.ToInt32(FrontsOrdersRow["InsetColorID"]);
                DataRow[] Rows = DecorConfigDataTable.Select("DecorID = " + InsetColorID);
                if (Rows.Count() > 0)
                {
                    DataRow[] CPRow = ClientsDataTable.Select("ClientID = " + ClientID);
                    if (CPRow.Count() > 0)
                    {
                        PriceGroup = Convert.ToDecimal(CPRow[0]["PriceGroup"]);//MarketingCost
                        MarketingCost = Convert.ToDecimal(Rows[0]["MarketingCost"]);
                        PriceRatio = Convert.ToDecimal(Rows[0]["PriceRatio"]);
                        DigitCapacity = Convert.ToInt32(Rows[0]["DigitCapacity"]);
                        OriginalPrice = MarketingCost * PriceRatio * (1 + PriceGroup);
                        OriginalPrice = Decimal.Round(OriginalPrice, DigitCapacity, MidpointRounding.AwayFromZero);

                        if (ClientID == 145)
                            OriginalPrice = Convert.ToDecimal(Rows[0]["ZOVPrice"]);
                        if (Sale)
                        {
                            decimal SampleSaleValue = GetSampleDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                            decimal SaleValue = GetDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                            if (IsSample)
                                Price = (OriginalPrice - OriginalPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                            else
                                Price = (OriginalPrice - OriginalPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                        }
                        else
                        {
                            if (IsSample)
                                Price = (OriginalPrice - OriginalPrice / 100 * 50) * (1 - DiscountDirector / 100);
                            else
                                Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                        }
                        //if (IsSample)
                        //{
                        //    Price = (OriginalPrice - OriginalPrice * 50 / 100) * (1 - DiscountDirector / 100);
                        //    //Price = OriginalPrice - OriginalPrice / 100 * (50 + DiscountDirector);
                        //}
                        //else
                        //{
                        //    Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                        //    //Price = CurrentPrice - CurrentPrice / 100 * (TotalDiscount + DiscountDirector);
                        //}
                        if (TechDrilling)
                            Price = Price * 1.3m;
                        FrontsOrdersRow["OriginalInsetPrice"] = OriginalPrice;
                        FrontsOrdersRow["InsetPrice"] = Price;
                        return;
                    }
                }
            }
            //Аппликации
            if (FrontID == 3728 || FrontID == 3731 ||
                FrontID == 3732 || FrontID == 3739 ||
                FrontID == 3740 || FrontID == 3741 ||
                FrontID == 3744 || FrontID == 3745 ||
                FrontID == 3746 ||
                InsetTypeID == 28961 || InsetTypeID == 3653 || InsetTypeID == 3654 || InsetTypeID == 3655
                )
            {
                OriginalPrice = 5;
                if (Sale)
                {
                    decimal SampleSaleValue = GetSampleDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                    decimal SaleValue = GetDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                    if (IsSample)
                        Price = (OriginalPrice - OriginalPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                    else
                        Price = (OriginalPrice - OriginalPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                }
                else
                {
                    if (IsSample)
                        Price = (OriginalPrice - OriginalPrice / 100 * 50) * (1 - DiscountDirector / 100);
                    else
                        Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                }
                if (TechDrilling)
                    Price = Price * 1.3m;
                //if (IsSample)
                //{
                //    Price = (OriginalPrice - OriginalPrice * 50 / 100) * (1 - DiscountDirector / 100);
                //    //Price = OriginalPrice - OriginalPrice / 100 * (50 + DiscountDirector);
                //}
                //else
                //{
                //    Price = (OriginalPrice - OriginalPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                //    //Price = CurrentPrice - CurrentPrice / 100 * (TotalDiscount + DiscountDirector);
                //}
            }
            FrontsOrdersRow["OriginalInsetPrice"] = OriginalPrice;
            FrontsOrdersRow["InsetPrice"] = Price;
            return;
        }

        private decimal GetNonStandardMargin(int FrontConfigID, int ClientID)
        {
            DataRow[] Rows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID);
            if (Rows.Count() == 0)
                return 0;
            else
                return Convert.ToDecimal(Rows[0]["ZOVNonStandMargin"]);
        }

        private bool IsMeasureSquare(DataRow FrontsOrdersRow)
        {
            DataRow[] Rows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());

            if (Rows.Count() > 0 && Convert.ToInt32(Rows[0]["MeasureID"]) == 1)
                return true;

            return false;
        }

        private bool IsInsetMeasureSquare(DataRow FrontsOrdersRow)
        {
            DataRow[] Rows = InsetTypesDataTable.Select("InsetTypeID = " + FrontsOrdersRow["InsetTypeID"].ToString());

            if (Rows.Count() > 0 && Convert.ToInt32(Rows[0]["MeasureID"]) == 1)
                return true;

            return false;
        }

        //ALUMINIUM
        public int IsAluminium(DataRow FrontsOrdersRow)
        {
            string str = FrontsOrdersRow["FrontID"].ToString();
            DataRow[] Row = FrontsDataTable.Select("FrontID = " + FrontsOrdersRow["FrontID"].ToString());

            if (Row.Count() > 0 && Row[0]["FrontName"].ToString()[0] == 'Z')
            {
                str = Row[0]["FrontName"].ToString();
                return Convert.ToInt32(Row[0]["FrontID"]);
            }

            return -1;
        }

        private void GetGlassMarginAluminium(DataRow FrontsOrdersRow, ref int GlassMarginHeight, ref int GlassMarginWidth)
        {
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + Convert.ToInt32(FrontsOrdersRow["FrontID"]));
            if (Rows.Count() > 0)
            {
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    GlassMarginHeight = Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    GlassMarginWidth = Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);
            }
            //DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(FrontsOrdersRow));
            //if (Rows.Count() > 0)
            //{
            //    GlassMarginHeight = Convert.ToInt32(Rows[0]["GlassMarginHeight"]);
            //    GlassMarginWidth = Convert.ToInt32(Rows[0]["GlassMarginWidth"]);
            //}
        }

        private decimal GetJobPriceAluminium(DataRow FrontsOrdersRow)
        {
            DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(FrontsOrdersRow));
            if (Rows.Count() == 0)
                return 0;
            return Convert.ToDecimal(Rows[0]["JobPrice"]);
        }

        private decimal FrontsPriceAluminium(DataRow FrontsOrdersRow)
        {
            decimal Price = 0;

            DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(FrontsOrdersRow));
            if (Rows.Count() > 0)
            {
                Price = Convert.ToDecimal(Rows[0]["ProfilPrice"]);
            }
            if (NeedCalcPrice)
                FrontsOrdersRow["FrontPrice"] = Price;

            return Price;
        }

        private decimal InsetPriceAluminium(DataRow FrontsOrdersRow)
        {
            decimal Price = 0;

            DataRow[] Rows = InsetPriceDataTable.Select("InsetTypeID = " + FrontsOrdersRow["InsetColorID"].ToString());

            if (Rows.Count() > 0)
                Price = Convert.ToDecimal(Rows[0]["GlassZXPrice"]);
            else
                Price = 0;
            if (NeedCalcPrice)
                FrontsOrdersRow["InsetPrice"] = Price;

            return Price;
        }

        public decimal GetFrontCostAluminium(int ClientID, DataRow FrontsOrdersRow)
        {
            decimal Cost = 0;
            decimal Perimeter = 0;
            decimal GlassSquare = 0;

            int MarginHeight = 0;
            int MarginWidth = 0;

            decimal GlassPrice = InsetPriceAluminium(FrontsOrdersRow);
            decimal JobPrice = GetJobPriceAluminium(FrontsOrdersRow);
            decimal ProfilPrice = FrontsPriceAluminium(FrontsOrdersRow);

            GetGlassMarginAluminium(FrontsOrdersRow, ref MarginHeight, ref MarginWidth);

            decimal Height = Convert.ToInt32(FrontsOrdersRow["Height"]);
            decimal Width = Convert.ToInt32(FrontsOrdersRow["Width"]);
            decimal Count = Convert.ToInt32(FrontsOrdersRow["Count"]);

            GlassPrice = Decimal.Round(GlassPrice, 2, MidpointRounding.AwayFromZero);
            JobPrice = Decimal.Round(JobPrice, 2, MidpointRounding.AwayFromZero);
            ProfilPrice = Decimal.Round(ProfilPrice, 2, MidpointRounding.AwayFromZero);

            Perimeter = Decimal.Round((Height * 2 + Width * 2) / 1000 * Count, 3, MidpointRounding.AwayFromZero);
            GlassSquare = Decimal.Round((Height - MarginHeight) * (Width - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);

            GlassSquare = GlassSquare * Count;
            Cost = Decimal.Round(JobPrice * Count + GlassPrice * GlassSquare + Perimeter * ProfilPrice, 3, MidpointRounding.AwayFromZero);

            if (ClientID != 145)
                Cost = Cost * 100 / 120;

            decimal d = Height * Width / 1000000;
            d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
            decimal Square = d * Count;

            Cost = Math.Ceiling(Cost / 0.01m) * 0.01m;
            FrontsOrdersRow["InsetPrice"] = 0;
            FrontsOrdersRow["Cost"] = Cost;
            FrontsOrdersRow["Square"] = Square;
            FrontsOrdersRow["FrontPrice"] = Cost / Square;

            return Cost;
        }

        public decimal GetFrontCostAluminium(int ClientID, DataRow FrontsOrdersRow, decimal DiscountDirector, decimal TotalDiscount, bool TechDrilling, ref decimal Square)
        {
            decimal Cost = 0;
            decimal Perimeter = 0;
            decimal GlassSquare = 0;

            int MarginHeight = 0;
            int MarginWidth = 0;

            decimal GlassPrice = InsetPriceAluminium(FrontsOrdersRow);
            decimal JobPrice = GetJobPriceAluminium(FrontsOrdersRow);
            decimal ProfilPrice = FrontsPriceAluminium(FrontsOrdersRow);

            GetGlassMarginAluminium(FrontsOrdersRow, ref MarginHeight, ref MarginWidth);

            decimal Height = Convert.ToInt32(FrontsOrdersRow["Height"]);
            decimal Width = Convert.ToInt32(FrontsOrdersRow["Width"]);
            decimal Count = Convert.ToInt32(FrontsOrdersRow["Count"]);

            bool IsSample = Convert.ToBoolean(FrontsOrdersRow["IsSample"]);//Образец

            if (Sale)
            {
                decimal SampleSaleValue = GetSampleDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                decimal SaleValue = GetDiscount(Convert.ToInt32(FrontsOrdersRow["FrontConfigID"]));
                if (IsSample)
                {
                    GlassPrice = (GlassPrice - GlassPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                    JobPrice = (JobPrice - JobPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                    ProfilPrice = (ProfilPrice - ProfilPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                }
                else
                {
                    GlassPrice = (GlassPrice - GlassPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                    JobPrice = (JobPrice - JobPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                    ProfilPrice = (ProfilPrice - ProfilPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                }
            }
            else
            {
                if (IsSample)
                {
                    GlassPrice = (GlassPrice - GlassPrice / 100 * 50) * (1 - DiscountDirector / 100);
                    JobPrice = (JobPrice - JobPrice / 100 * 50) * (1 - DiscountDirector / 100);
                    ProfilPrice = (ProfilPrice - ProfilPrice / 100 * 50) * (1 - DiscountDirector / 100);
                }
                else
                {
                    GlassPrice = (GlassPrice - GlassPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    JobPrice = (JobPrice - JobPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    ProfilPrice = (ProfilPrice - ProfilPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                }
            }
            //if (IsSample)
            //{
            //    GlassPrice = (GlassPrice - GlassPrice * 50 / 100) * (1 - DiscountDirector / 100);
            //    JobPrice = (JobPrice - JobPrice * 50 / 100) * (1 - DiscountDirector / 100);
            //    ProfilPrice = (ProfilPrice - ProfilPrice * 50 / 100) * (1 - DiscountDirector / 100);
            //}
            //else
            //{
            //    GlassPrice = (GlassPrice - GlassPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
            //    JobPrice = (JobPrice - JobPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
            //    ProfilPrice = (ProfilPrice - ProfilPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
            //}
            GlassPrice = Decimal.Round(GlassPrice, 2, MidpointRounding.AwayFromZero);
            JobPrice = Decimal.Round(JobPrice, 2, MidpointRounding.AwayFromZero);
            ProfilPrice = Decimal.Round(ProfilPrice, 2, MidpointRounding.AwayFromZero);

            Perimeter = Decimal.Round((Height * 2 + Width * 2) / 1000 * Count, 3, MidpointRounding.AwayFromZero);
            GlassSquare = Decimal.Round((Height - MarginHeight) * (Width - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);

            GlassSquare = GlassSquare * Count;
            Cost = Decimal.Round(JobPrice * Count + GlassPrice * GlassSquare + Perimeter * ProfilPrice, 3, MidpointRounding.AwayFromZero);

            if (ClientID != 145)
                Cost = Cost * 100 / 120;

            decimal d = Height * Width / 1000000;
            d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
            Square = d * Count;

            if (TechDrilling)
                Cost = Cost * 1.3m;
            Cost = Math.Ceiling(Cost / 0.01m) * 0.01m;
            FrontsOrdersRow["InsetPrice"] = 0;
            FrontsOrdersRow["Cost"] = Cost;
            FrontsOrdersRow["Square"] = Square;
            if (NeedCalcPrice)
                FrontsOrdersRow["FrontPrice"] = Cost / Square;

            return Cost;
        }

        private void ProfilCalculateItem(int ClientID, DataRow FrontsOrdersRow, decimal DiscountDirector, decimal TotalDiscount,
            int CurrencyTypeID, decimal Currency, bool TechDrilling,
            ref decimal OriginalCost, ref decimal Cost, ref decimal CurrencyOrderCost, ref decimal ItemSquare)
        {
            int FrontID = Convert.ToInt32(FrontsOrdersRow["FrontID"]);
            int InsetTypeID = Convert.ToInt32(FrontsOrdersRow["InsetTypeID"]);
            decimal Height = Convert.ToDecimal(FrontsOrdersRow["Height"]);
            decimal Width = Convert.ToDecimal(FrontsOrdersRow["Width"]);
            decimal Count = Convert.ToDecimal(FrontsOrdersRow["Count"]);

            decimal Square = 0;

            decimal OriginalFrontCost = 0;
            decimal FrontCost = 0;
            decimal InsetCost = 0;

            //ALUMINIUM
            if (IsAluminium(FrontsOrdersRow) > -1)
            {
                Cost = GetFrontCostAluminium(ClientID, FrontsOrdersRow, DiscountDirector, TotalDiscount, TechDrilling, ref Square);

                decimal OriginalFrontPrice = Cost / Square;
                decimal dFrontPrice = Cost / Square;
                OriginalFrontCost = Square * OriginalFrontPrice;
                FrontCost = Square * dFrontPrice;

                decimal d = dFrontPrice * Currency;
                //if (CurrencyTypeID == 4)
                //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                CurrencyOrderCost = Square * d;

                OriginalCost = Cost;
                //Cost = Cost * Currency;
            }
            else
            {
                decimal OriginalFrontPrice = 0;
                decimal dFrontPrice = 0;
                decimal OriginalInsetPrice = 0;
                decimal dInsetPrice = 0;
                if (NeedCalcPrice)
                {
                    ProfilFrontPrice(ClientID, FrontsOrdersRow, DiscountDirector, TotalDiscount, TechDrilling, 
                        ref OriginalFrontPrice, ref dFrontPrice, Currency);
                    InsetPrice(ClientID, FrontsOrdersRow, DiscountDirector, TotalDiscount, TechDrilling, ref OriginalInsetPrice, ref dInsetPrice);
                }
                else
                {
                    OriginalFrontPrice = Convert.ToDecimal(FrontsOrdersRow["OriginalPrice"]);
                    dFrontPrice = Convert.ToDecimal(FrontsOrdersRow["FrontPrice"]);
                    OriginalInsetPrice = Convert.ToDecimal(FrontsOrdersRow["OriginalInsetPrice"]);
                    dInsetPrice = Convert.ToDecimal(FrontsOrdersRow["InsetPrice"]);
                }

                if (IsMeasureSquare(FrontsOrdersRow))
                {
                    decimal d = Height * Width / 1000000;
                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                    Square = d * Count;
                    OriginalFrontCost = Square * OriginalFrontPrice;
                    FrontCost = Square * dFrontPrice;

                    d = dFrontPrice * Currency;
                    //if (CurrencyTypeID == 4)
                    //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                    CurrencyOrderCost = Square * d;
                }
                else
                {
                    OriginalFrontCost = OriginalFrontPrice * Count;
                    FrontCost = dFrontPrice * Count;

                    decimal d = dFrontPrice * Currency;
                    //if (CurrencyTypeID == 4)
                    //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                    CurrencyOrderCost = Count * d;
                }

                if (dInsetPrice > 0)
                {
                    if (FrontID == 3728 || FrontID == 3731 ||
                        FrontID == 3732 || FrontID == 3739 ||
                        FrontID == 3740 || FrontID == 3741 ||
                        FrontID == 3744 || FrontID == 3745 ||
                        FrontID == 3746 ||
                        InsetTypeID == 28961 || InsetTypeID == 3653 || InsetTypeID == 3654 || InsetTypeID == 3655
                        )
                    {
                        InsetCost = dInsetPrice * Count;

                        decimal d = dInsetPrice * Currency;
                        //if (CurrencyTypeID == 4)
                        //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                        CurrencyOrderCost += Count * d;
                    }
                    else
                    {
                        decimal dInsetSquare = GetInsetSquare(FrontsOrdersRow);
                        InsetCost = dInsetPrice * dInsetSquare * Count;

                        decimal d = dInsetPrice * Currency;
                        //if (CurrencyTypeID == 4)
                        //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                        CurrencyOrderCost += dInsetSquare * Count * d;
                    }
                }
                FrontsOrdersRow["InsetPrice"] = dInsetPrice;
            }

            FrontCost = Math.Ceiling(FrontCost / 0.01m) * 0.01m;
            InsetCost = Math.Ceiling(InsetCost / 0.01m) * 0.01m;
            OriginalFrontCost = Math.Ceiling(OriginalFrontCost / 0.01m) * 0.01m;
            OriginalCost = OriginalFrontCost + InsetCost;
            Cost = FrontCost + InsetCost;

            FrontsOrdersRow["Square"] = Square;
            FrontsOrdersRow["OriginalCost"] = OriginalCost;
            FrontsOrdersRow["CurrencyTypeID"] = CurrencyTypeID;
            FrontsOrdersRow["Cost"] = Cost;
            CurrencyOrderCost = Math.Ceiling(CurrencyOrderCost / 0.01m) * 0.01m;
            //CurrencyOrderCost = Decimal.Round(CurrencyOrderCost, 2, MidpointRounding.AwayFromZero);
            FrontsOrdersRow["CurrencyCost"] = CurrencyOrderCost;

            ItemSquare = Square;
        }

        private void TPSCalculateItem(int ClientID, DataRow FrontsOrdersRow, decimal DiscountDirector, decimal TotalDiscount,
            int CurrencyTypeID, decimal Currency, bool TechDrilling,
            ref decimal OriginalCost, ref decimal Cost, ref decimal CurrencyOrderCost, ref decimal ItemSquare)
        {
            int FrontID = Convert.ToInt32(FrontsOrdersRow["FrontID"]);
            int InsetTypeID = Convert.ToInt32(FrontsOrdersRow["InsetTypeID"]);
            decimal Height = Convert.ToDecimal(FrontsOrdersRow["Height"]);
            decimal Width = Convert.ToDecimal(FrontsOrdersRow["Width"]);
            decimal Count = Convert.ToDecimal(FrontsOrdersRow["Count"]);
            decimal Square = 0;

            decimal OriginalFrontCost = 0;
            decimal FrontCost = 0;
            decimal InsetCost = 0;
            decimal FrontCost1 = 0;
            decimal InsetCost1 = 0;

            //ALUMINIUM
            if (IsAluminium(FrontsOrdersRow) > -1)
            {
                Cost = GetFrontCostAluminium(ClientID, FrontsOrdersRow, DiscountDirector, TotalDiscount, TechDrilling, ref Square);
                decimal OriginalFrontPrice = Cost / Square;
                decimal dFrontPrice = Cost / Square;
                OriginalFrontCost = Square * OriginalFrontPrice;
                FrontCost = Square * dFrontPrice;

                decimal d = dFrontPrice * Currency;
                //if (CurrencyTypeID == 4)
                //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                CurrencyOrderCost = Square * d;

                OriginalCost = Cost;
                //Cost = Cost * Currency;
            }
            else
            {
                decimal OriginalFrontPrice = 0;
                decimal dFrontPrice = 0;
                decimal OriginalInsetPrice = 0;
                decimal dInsetPrice = 0;
                if (NeedCalcPrice)
                {
                    TPSFrontPrice(ClientID, FrontsOrdersRow, DiscountDirector, TotalDiscount, TechDrilling, ref OriginalFrontPrice, ref dFrontPrice);
                    InsetPrice(ClientID, FrontsOrdersRow, DiscountDirector, TotalDiscount, TechDrilling, ref OriginalInsetPrice, ref dInsetPrice);
                }
                else
                {
                    OriginalFrontPrice = Convert.ToDecimal(FrontsOrdersRow["OriginalPrice"]);
                    dFrontPrice = Convert.ToDecimal(FrontsOrdersRow["FrontPrice"]);
                    dInsetPrice = Convert.ToDecimal(FrontsOrdersRow["InsetPrice"]);
                }

                if (IsMeasureSquare(FrontsOrdersRow))
                {
                    //Square = Height * Width / 1000000 * Count;
                    decimal d = Height * Width / 1000000;
                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                    Square = d * Count;
                    OriginalFrontCost = Square * OriginalFrontPrice;
                    FrontCost = Square * dFrontPrice;

                    d = dFrontPrice * Currency;
                    //if (CurrencyTypeID == 4)
                    //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                    FrontCost1 = Square * d;
                }
                else
                {
                    OriginalFrontCost = OriginalFrontPrice * Count;
                    FrontCost = dFrontPrice * Count;

                    decimal d = dFrontPrice * Currency;
                    //if (CurrencyTypeID == 4)
                    //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                    FrontCost1 = Count * d;
                }

                if (dInsetPrice > 0)
                {
                    if (FrontID == 3728 || FrontID == 3731 ||
                        FrontID == 3732 || FrontID == 3739 ||
                        FrontID == 3740 || FrontID == 3741 ||
                        FrontID == 3744 || FrontID == 3745 ||
                        FrontID == 3746 ||
                        InsetTypeID == 28961 || InsetTypeID == 3653 || InsetTypeID == 3654 || InsetTypeID == 3655
                        )
                    {
                        InsetCost = dInsetPrice * Count;

                        decimal d = dInsetPrice * Currency;
                        //if (CurrencyTypeID == 4)
                        //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                        InsetCost1 = Count * d;
                    }
                    else
                    {
                        decimal dInsetSquare = GetInsetSquare(FrontsOrdersRow);
                        if (FrontID == 3729)
                        {
                            int MarginHeight = 0;
                            int MarginWidth = 0;
                            GetGlassMarginAluminium(FrontsOrdersRow, ref MarginHeight, ref MarginWidth);
                            dInsetSquare = Decimal.Round(MarginHeight * (Convert.ToDecimal(FrontsOrdersRow["Width"]) - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);
                        }
                        InsetCost = dInsetPrice * dInsetSquare * Count;

                        decimal d = dInsetPrice * Currency;
                        //if (CurrencyTypeID == 4)
                        //    d = Math.Ceiling(d / 100.0m) * 100.0m;
                        InsetCost1 = dInsetSquare * Count * d;
                    }
                }
                FrontsOrdersRow["InsetPrice"] = dInsetPrice;

                FrontCost = Math.Ceiling(FrontCost / 0.01m) * 0.01m;
                InsetCost = Math.Ceiling(InsetCost / 0.01m) * 0.01m;
                FrontCost1 = Math.Ceiling(FrontCost1 / 0.01m) * 0.01m;
                InsetCost1 = Math.Ceiling(InsetCost1 / 0.01m) * 0.01m;
                OriginalFrontCost = Math.Ceiling(OriginalFrontCost / 0.01m) * 0.01m;
                OriginalCost = OriginalFrontCost + InsetCost;
                Cost = FrontCost + InsetCost;
                CurrencyOrderCost = FrontCost1 + InsetCost1;
            }

            FrontsOrdersRow["Square"] = Square;
            FrontsOrdersRow["OriginalCost"] = OriginalCost;
            FrontsOrdersRow["CurrencyTypeID"] = CurrencyTypeID;
            FrontsOrdersRow["Cost"] = Cost;
            CurrencyOrderCost = Math.Ceiling(CurrencyOrderCost / 0.01m) * 0.01m;
            //CurrencyOrderCost = Decimal.Round(CurrencyOrderCost, 2, MidpointRounding.AwayFromZero);
            FrontsOrdersRow["CurrencyCost"] = CurrencyOrderCost;

            ItemSquare = Square;
        }

        private decimal GetInsetSquare(DataRow FrontsOrdersRow)
        {
            decimal GridHeight = 0;
            decimal GridWidth = 0;
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + Convert.ToInt32(FrontsOrdersRow["FrontID"]));
            if (Rows.Count() > 0)
            {
                decimal InsetHeightAdmission = Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                decimal InsetWidthAdmission = Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    GridHeight = Convert.ToInt32(FrontsOrdersRow["Height"]) - Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    GridWidth = Convert.ToInt32(FrontsOrdersRow["Width"]) - Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);
                if (Convert.ToInt32(FrontsOrdersRow["FrontID"]) == 3729)
                {
                    return Decimal.Round(Convert.ToInt32(Rows[0]["InsetHeightAdmission"]) * GridWidth / 1000000, 3, MidpointRounding.AwayFromZero);
                }
            }
            return Decimal.Round(GridHeight * GridWidth / 1000000, 3, MidpointRounding.AwayFromZero);
        }

        private decimal GetInsetWeight(DataRow FrontsOrdersRow)
        {
            int InsetTypeID = Convert.ToInt32(FrontsOrdersRow["InsetTypeID"]);
            if (InsetTypeID == 2)
                InsetTypeID = Convert.ToInt32(FrontsOrdersRow["InsetColorID"]);//стекло
            decimal InsetSquare = GetInsetSquare(FrontsOrdersRow);
            if (InsetSquare <= 0)
                return 0;
            decimal InsetWeight = 0;
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + InsetTypeID);
            if (Rows.Count() > 0)
            {
                if (Rows[0]["Weight"] != DBNull.Value)
                    InsetWeight = Convert.ToDecimal(Rows[0]["Weight"]);
            }

            return InsetSquare * InsetWeight;
        }

        private decimal GetProfileWeight(DataRow FrontsOrdersRow)
        {
            decimal FrontHeight = Convert.ToDecimal(FrontsOrdersRow["Height"]);
            decimal FrontWidth = Convert.ToDecimal(FrontsOrdersRow["Width"]);
            decimal ProfileWeight = 0;
            decimal ProfileWidth = 0;
            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());
            if (FrontsConfigRow.Count() > 0)
                ProfileWeight = Convert.ToDecimal(FrontsConfigRow[0]["Weight"]);

            int FrontID = Convert.ToInt32(FrontsOrdersRow["FrontID"]);
            if (Security.IsFrontsSquareCalc(FrontID))
            {
                return FrontWidth * FrontHeight / 1000000 * ProfileWeight;
            }
            //if (FrontID == 30504 || FrontID == 30505 || FrontID == 30506 ||
            //    FrontID == 30364 || FrontID == 30366 || FrontID == 30367 ||
            //    FrontID == 30501 || FrontID == 30502 || FrontID == 30503 ||
            //    FrontID == 16269 || FrontID == 28945 || FrontID == 41327 || FrontID == 41328 || FrontID == 41331 || 
            //    FrontID == 27914 || FrontID == 29597 || FrontID == 3727 || FrontID == 3728 || FrontID == 3729 ||
            //    FrontID == 3730 || FrontID == 3731 || FrontID == 3732 || FrontID == 3733 || FrontID == 3734 ||
            //    FrontID == 3735 || FrontID == 3736 || FrontID == 3737 || FrontID == 3739 || FrontID == 3740 ||
            //    FrontID == 3741 || FrontID == 3742 || FrontID == 3743 || FrontID == 3744 || FrontID == 3745 ||
            //    FrontID == 3746 || FrontID == 3747 || FrontID == 3748 || FrontID == 15108 || FrontID == 3662 || FrontID == 3663 || FrontID == 3664 || FrontID == 15760)
            //    return FrontWidth * FrontHeight / 1000000 * ProfileWeight;
            else
            {
                DataRow[] DecorConfigRow = TechStoreDataTable.Select("TechStoreID = " + FrontsConfigRow[0]["ProfileID"].ToString());
                if (DecorConfigRow.Count() > 0)
                {
                    ProfileWidth = Convert.ToDecimal(DecorConfigRow[0]["Width"]);
                    ProfileWeight = Convert.ToDecimal(DecorConfigRow[0]["Weight"]);
                    return (FrontWidth * 2 + (FrontHeight - ProfileWidth - ProfileWidth) * 2) / 1000 * ProfileWeight;
                }
            }
            return 0;
        }

        public decimal GetAluminiumWeight(DataRow FrontsOrdersRow, bool WithGlass)
        {
            DataRow[] Row = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(FrontsOrdersRow));
            if (Row.Count() == 0)
                return 0;
            decimal FrontHeight = Convert.ToDecimal(FrontsOrdersRow["Height"]);
            decimal FrontWidth = Convert.ToDecimal(FrontsOrdersRow["Width"]);
            decimal Count = Convert.ToInt32(FrontsOrdersRow["Count"]);

            int MarginHeight = 0;
            int MarginWidth = 0;

            GetGlassMarginAluminium(FrontsOrdersRow, ref MarginHeight, ref MarginWidth);

            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());
            if (FrontsConfigRow.Count() == 0)
                return 0;
            int ProfileID = Convert.ToInt32(FrontsConfigRow[0]["ProfileID"]);
            decimal ProfileWeight = 0;
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + ProfileID);
            if (Rows.Count() > 0)
            {
                if (Rows[0]["Weight"] != DBNull.Value)
                    ProfileWeight = Convert.ToDecimal(Rows[0]["Weight"]);
            }

            decimal GlassSquare = 0;

            if (FrontsOrdersRow["InsetColorID"].ToString() != "3946")//если не СТЕКЛО КЛИЕНТА
                GlassSquare = Decimal.Round((FrontHeight - MarginHeight) * (FrontWidth - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);

            decimal GlassWeight = GlassSquare * 10;

            decimal ResultProfileWeight = Decimal.Round((FrontWidth * 2 + FrontHeight * 2) / 1000, 3, MidpointRounding.AwayFromZero) * ProfileWeight;

            if (WithGlass)
                return (ResultProfileWeight + GlassWeight) * Count;
            else
                return (ResultProfileWeight) * Count;
        }

        public decimal CalculateFrontsWeight(DataRow FrontsOrdersRow)
        {
            decimal FrontsWeight = 0;
            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());
            decimal InsetWeight = Convert.ToDecimal(FrontsConfigRow[0]["InsetWeight"]);
            //если гнутый то вес за штуки
            if (FrontsConfigRow[0]["Width"].ToString() == "-1")
            {
                FrontsOrdersRow["ItemWeight"] = FrontsConfigRow[0]["Weight"];
                FrontsOrdersRow["Weight"] = Convert.ToDecimal(FrontsOrdersRow["Count"]) * Convert.ToDecimal(FrontsConfigRow[0]["Weight"]);
                return Convert.ToDecimal(FrontsOrdersRow["Count"]) * Convert.ToDecimal(FrontsConfigRow[0]["Weight"]);
            }
            //если алюминий
            if (IsAluminium(FrontsOrdersRow) > -1)
            {
                decimal W = GetAluminiumWeight(FrontsOrdersRow, true);
                FrontsOrdersRow["ItemWeight"] = W / Convert.ToDecimal(FrontsOrdersRow["Count"]);
                FrontsOrdersRow["Weight"] = W;
                return W;
            }
            decimal ResultProfileWeight = GetProfileWeight(FrontsOrdersRow);
            decimal ResultInsetWeight = 0;
            DataRow[] rows = InsetTypesDataTable.Select("GroupID = 2 OR GroupID = 3 OR GroupID = 4 OR GroupID = 7 OR GroupID = 12 OR GroupID = 13");
            foreach (DataRow item in rows)
            {
                if (FrontsOrdersRow["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                {
                    ResultInsetWeight = GetInsetWeight(FrontsOrdersRow);
                }
            }
            if (Convert.ToInt32(FrontsOrdersRow["FrontID"]) == 3729)
                ResultInsetWeight = GetInsetWeight(FrontsOrdersRow);

            FrontsWeight = ResultProfileWeight + ResultInsetWeight;
            FrontsWeight = Decimal.Round(FrontsWeight, 3, MidpointRounding.AwayFromZero);
            FrontsOrdersRow["ItemWeight"] = FrontsWeight;
            FrontsOrdersRow["Weight"] = FrontsWeight * Convert.ToInt32(FrontsOrdersRow["Count"]);

            return FrontsWeight * Convert.ToInt32(FrontsOrdersRow["Count"]);
        }

        //общий расчет фасадов в заказе
        public void CalculateFronts(int ClientID, int MainOrderID, decimal ProfilDiscountDirector, decimal TPSDiscountDirector, decimal ProfilTotalDiscount, decimal TPSTotalDiscount, int CurrencyTypeID, decimal Currency, bool TechDrilling,
            ref decimal OriginalProfilCost, ref decimal OriginalTPSCost, ref decimal OriginalTotalCost,
            ref decimal ProfilCost, ref decimal TPSCost, ref decimal TotalCost, ref decimal TotalCurrencyCost,
            ref decimal OrderSquare, ref decimal FrontsWeight, object ConfirmDateTime)
        {
            GetStocks(ClientID, ConfirmDateTime);
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders WHERE MainOrderID = " + MainOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable FrontsOrdersDataTable = new DataTable())
                    {
                        DA.Fill(FrontsOrdersDataTable);

                        if (FrontsOrdersDataTable.Rows.Count == 0)
                            return;

                        foreach (DataRow Row in FrontsOrdersDataTable.Rows)
                        {
                            decimal ItemSquare = 0;
                            decimal OriginalCost = 0;
                            decimal Cost = 0;
                            decimal CurrencyCost = 0;

                            int FactoryID = Convert.ToInt32(Row["FactoryID"]);
                            int FrontsOrdersID = Convert.ToInt32(Row["FrontsOrdersID"]);

                            FrontsWeight += CalculateFrontsWeight(Row);
                            if (FactoryID == 1)
                            {
                                ProfilCalculateItem(ClientID, Row, ProfilDiscountDirector, ProfilTotalDiscount, CurrencyTypeID, Currency, TechDrilling,
                                    ref OriginalCost, ref Cost, ref CurrencyCost, ref ItemSquare);
                                OriginalProfilCost += OriginalCost;
                                ProfilCost += Cost;
                            }
                            if (FactoryID == 2)
                            {
                                TPSCalculateItem(ClientID, Row, TPSDiscountDirector, TPSTotalDiscount, CurrencyTypeID, Currency, TechDrilling,
                                    ref OriginalCost, ref Cost, ref CurrencyCost, ref ItemSquare);
                                OriginalTPSCost += OriginalCost;
                                TPSCost += Cost;
                            }
                            TotalCurrencyCost += CurrencyCost;
                            OrderSquare += ItemSquare;
                        }
                        OriginalTotalCost = OriginalProfilCost + OriginalTPSCost;
                        TotalCost = ProfilCost + TPSCost;

                        DA.Update(FrontsOrdersDataTable);
                    }
                }
            }
        }

        //общий расчет фасадов в заказе
        public void AssignTransportCost(int MegaOrderID, decimal ProfilComplaintCost, decimal TPSComplaintCost, decimal TransportAdditionalCost)
        {
            decimal ProfilTotalWeight = 0;
            decimal TPSTotalWeight = 0;
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow Row in DT.Rows)
                            {
                                decimal Square = Convert.ToDecimal(Row["Square"]);
                                decimal Weight = Convert.ToDecimal(Row["Weight"]);
                                if (Square > 0)
                                    Weight = Weight + Square * Convert.ToDecimal(0.7);
                                Weight = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
                                ProfilTotalWeight += Weight;
                            }
                        }
                    }
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorOrders WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow Row in DT.Rows)
                                ProfilTotalWeight += Convert.ToDecimal(Row["Weight"]);
                        }
                    }
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow Row in DT.Rows)
                            {
                                decimal Square = Convert.ToDecimal(Row["Square"]);
                                decimal Weight = Convert.ToDecimal(Row["Weight"]);
                                if (Square > 0)
                                    Weight = Weight + Square * Convert.ToDecimal(0.7);
                                Weight = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
                                TPSTotalWeight += Weight;
                            }
                        }
                    }
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorOrders WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow Row in DT.Rows)
                                TPSTotalWeight += Convert.ToDecimal(Row["Weight"]);
                        }
                    }
                }
            }
            decimal ProfilTransportAdditionalCost = 0;
            decimal TPSTransportAdditionalCost = 0;
            if (ProfilTotalWeight > 0)
                ProfilTransportAdditionalCost = ProfilComplaintCost + TransportAdditionalCost * ProfilTotalWeight / (ProfilTotalWeight + TPSTotalWeight);
            if (TPSTotalWeight > 0)
                TPSTransportAdditionalCost = TPSComplaintCost + TransportAdditionalCost * TPSTotalWeight / (ProfilTotalWeight + TPSTotalWeight);
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders WHERE MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable FrontsOrdersDataTable = new DataTable())
                    {
                        DA.Fill(FrontsOrdersDataTable);

                        if (FrontsOrdersDataTable.Rows.Count == 0)
                            return;

                        foreach (DataRow Row in FrontsOrdersDataTable.Rows)
                        {
                            decimal PriceWithTransport = 0;
                            decimal CostWithTransport = 0;
                            int Count = Convert.ToInt32(Row["Count"]);
                            int FactoryID = Convert.ToInt32(Row["FactoryID"]);
                            decimal Price = Convert.ToDecimal(Row["FrontPrice"]);
                            decimal Cost = Convert.ToDecimal(Row["Cost"]);
                            decimal Weight = Convert.ToDecimal(Row["Weight"]);
                            decimal Square = Convert.ToDecimal(Row["Square"]);

                            decimal d = 0;
                            decimal d1 = 0;
                            if (Square > 0)
                                Weight = Weight + Square * Convert.ToDecimal(0.7);
                            Weight = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);

                            if (FactoryID == 1)
                            {
                                if (ProfilTotalWeight == 0)
                                    continue;
                                d = Weight / (ProfilTotalWeight / 100);
                                d1 = ProfilTransportAdditionalCost / 100 * d;
                            }
                            if (FactoryID == 2)
                            {
                                if (TPSTotalWeight == 0)
                                    continue;
                                d = Weight / (TPSTotalWeight / 100);
                                d1 = TPSTransportAdditionalCost / 100 * d;
                            }
                            if (Square > 0)
                            {
                                if (Square == 0)
                                {
                                    System.Windows.Forms.MessageBox.Show("Кол-во продукции не может быть равно 0");
                                }
                                else
                                {
                                    PriceWithTransport = Price + d1 / Count;
                                }
                            }
                            else
                            {
                                if (Count == 0)
                                {
                                    System.Windows.Forms.MessageBox.Show("Кол-во продукции не может быть равно 0");
                                }
                                else
                                {
                                    PriceWithTransport = Price + d1 / Count;
                                }
                            }
                            PriceWithTransport = Decimal.Round(PriceWithTransport, 2, MidpointRounding.AwayFromZero);
                            Row["PriceWithTransport"] = PriceWithTransport;
                            CostWithTransport = Cost + d1;
                            CostWithTransport = Decimal.Round(CostWithTransport, 2, MidpointRounding.AwayFromZero);
                            Row["CostWithTransport"] = CostWithTransport;
                        }

                        DA.Update(FrontsOrdersDataTable);
                    }
                }
            }
        }

        private void GetStocks(int ClientID, object ConfirmDateTime)
        {
            Sale = false;
            DateTime DocDateTime = DateTime.Now;
            if (ConfirmDateTime != DBNull.Value)
                DocDateTime = Convert.ToDateTime(ConfirmDateTime);

            if (StockDetailsDT == null)
                StockDetailsDT = new DataTable();
            else
                StockDetailsDT.Clear();

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT StockDetails.*, Stocks.Discount, Stocks.SampleDiscount FROM StockDetails INNER JOIN Stocks ON StockDetails.StockID=Stocks.StockID AND Stocks.Enable=1 AND 
                CAST(Stocks.FirstDate AS Date) <= '" + DocDateTime.ToString("yyyy-MM-dd") +
               "' AND CAST(Stocks.LastDate AS Date) >= '" + DocDateTime.ToString("yyyy-MM-dd") + "' WHERE StockDetails.StockID IN (SELECT StockID FROM StockClients WHERE ClientID=" + ClientID + ") AND StockDetails.ProductType=0",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                if (DA.Fill(StockDetailsDT) > 0)
                    Sale = true;
            }
        }

        private decimal GetDiscount(int ConfigID)
        {
            DataRow[] rows = StockDetailsDT.Select("ConfigID = " + ConfigID);
            if (rows.Count() > 0)
                return Convert.ToDecimal(rows[0]["Discount"]);
            return 0;
        }

        private decimal GetSampleDiscount(int ConfigID)
        {
            DataRow[] rows = StockDetailsDT.Select("ConfigID = " + ConfigID);
            if (rows.Count() > 0)
                return Convert.ToDecimal(rows[0]["SampleDiscount"]);
            return 0;
        }
    }






    public class DecorCalculate
    {
        private bool Sale = false;
        public bool NeedCalcPrice = false;

        private DataTable StockDetailsDT = null;
        private DataTable DecorConfigDataTable = null;
        private DataTable ClientsDataTable = null;
        private DataTable DiscountVolumeTable = null;
        private DataTable DecorOrdersTable = null;
        private DataTable AllDecorOrdersTable = null;

        private FrontsCalculate FrontsCalculate;

        public DecorCalculate(FrontsCalculate tFrontsCalculate)
        {
            FrontsCalculate = tFrontsCalculate;

            Initialize();
        }

        private void Fill()
        {
            DecorConfigDataTable = new DataTable();
            DecorConfigDataTable = TablesManager.DecorConfigDataTableAll;
            //using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT DecorConfig.*, Decor.TechStoreName FROM DecorConfig INNER JOIN
            //    infiniu2_catalog.dbo.TechStore AS Decor ON DecorConfig.DecorID = Decor.TechStoreID",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDataTable);
            //}

            ClientsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Clients",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsDataTable);
            }
            DiscountVolumeTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DiscountVolume",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(DiscountVolumeTable);
            }
            DecorOrdersTable = new DataTable();
            AllDecorOrdersTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 * FROM DecorOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersTable);
                DA.Fill(AllDecorOrdersTable);
            }
        }

        private void Initialize()
        {
            Fill();
        }

        public decimal GetDiscountVolume(decimal TotalVolume)
        {
            string withDot = TotalVolume.ToString(System.Globalization.CultureInfo.InvariantCulture);
            decimal Discount = 0;
            DataRow[] Rows = DiscountVolumeTable.Select("MinVolume < " + withDot + " AND MaxVolume >=" + withDot);
            if (Rows.Count() > 0)
                Discount = Convert.ToDecimal(Rows[0]["Discount"]);
            return Discount;
        }

        private void GetPilastrPrice(int clientId, DataRow decorOrderRow, decimal length, decimal discountDirector, decimal totalDiscount, ref decimal originalPrice, ref decimal price, decimal discountVolume = 1)
        {
            decimal currentPrice = 0;
            decimal marketingCost = 0;//себестоимость
            decimal priceRatio = 0;//ценовой коэффициент
            var digitCapacity = 0;//разрядность
            decimal priceGroup = 0;//ценовая группа

            var fpRow = ClientsDataTable.Select("ClientID = " + clientId);
            if (fpRow.Count() > 0)
            {
                var isSample = Convert.ToBoolean(decorOrderRow["IsSample"]);//Образец
                priceGroup = Convert.ToDecimal(fpRow[0]["PriceGroup"]);//MarketingCost
                var decorConfigId = Convert.ToInt32(decorOrderRow["DecorConfigID"]);
                var pRows = DecorConfigDataTable.Select("DecorConfigID = " + decorOrderRow["DecorConfigID"]);
                if (pRows.Count() > 0)
                {
                    marketingCost = Convert.ToDecimal(pRows[0]["MarketingCost"]);
                    priceRatio = Convert.ToDecimal(pRows[0]["PriceRatio"]);
                    digitCapacity = Convert.ToInt32(pRows[0]["DigitCapacity"]);
                    originalPrice = marketingCost * priceRatio * (1 + priceGroup);
                    originalPrice = decimal.Round(originalPrice, digitCapacity, MidpointRounding.AwayFromZero);
                    currentPrice = discountVolume * originalPrice;

                    var pPrice = 50 / (length / 1000);

                    currentPrice += pPrice;

                    if (clientId == 145)
                    {
                        originalPrice = Convert.ToDecimal(pRows[0]["ZOVPrice"]);
                        currentPrice = discountVolume * originalPrice;
                        currentPrice += 60 / (length / 1000);
                    }
                    price = currentPrice - currentPrice / 100 * totalDiscount;
                    if (Sale)
                    {
                        var sampleSaleValue = GetSampleDiscount(Convert.ToInt32(decorOrderRow["DecorConfigID"]));
                        var saleValue = GetDiscount(Convert.ToInt32(decorOrderRow["DecorConfigID"]));
                        if (isSample)
                            price = (currentPrice - currentPrice / 100 * (50 + sampleSaleValue)) * (1 - discountDirector / 100);
                        else
                            price = (currentPrice - currentPrice / 100 * (totalDiscount + saleValue)) * (1 - discountDirector / 100);
                    }
                    else
                    {
                        if (isSample)
                            price = (currentPrice - currentPrice / 100 * 50) * (1 - discountDirector / 100);
                        else
                            price = (currentPrice - currentPrice / 100 * totalDiscount) * (1 - discountDirector / 100);
                    }
                    //if (IsSample)
                    //{
                    //    Price = (CurrentPrice - CurrentPrice * 50 / 100) * (1 - DiscountDirector / 100);
                    //    //Price = OriginalPrice - OriginalPrice / 100 * (50 + DiscountDirector);
                    //}
                    //else
                    //{
                    //    Price = (CurrentPrice - CurrentPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    //    //Price = CurrentPrice - CurrentPrice / 100 * (TotalDiscount + DiscountDirector);
                    //}
                    price = decimal.Round(price, 2, MidpointRounding.AwayFromZero);
                }
            }
        }

        private void GetProfilDecorPrice(int ClientID, DataRow DecorOrderRow, decimal DiscountDirector, decimal TotalDiscount, ref decimal OriginalPrice, ref decimal Price, decimal DiscountVolume = 1)
        {
            decimal CurrentPrice = 0;
            decimal MarketingCost = 0;//себестоимость
            decimal PriceRatio = 0;//ценовой коэффициент
            int DigitCapacity = 0;//разрядность
            decimal PriceGroup = 0;//ценовая группа

            DataRow[] FPRow = ClientsDataTable.Select("ClientID = " + ClientID);
            if (FPRow.Count() > 0)
            {
                bool IsSample = Convert.ToBoolean(DecorOrderRow["IsSample"]);//Образец
                PriceGroup = Convert.ToDecimal(FPRow[0]["PriceGroup"]);//MarketingCost
                int DecorConfigID = Convert.ToInt32(DecorOrderRow["DecorConfigID"]);
                DataRow[] PRows = DecorConfigDataTable.Select("DecorConfigID = " + DecorOrderRow["DecorConfigID"].ToString());
                if (PRows.Count() > 0)
                {
                    MarketingCost = Convert.ToDecimal(PRows[0]["MarketingCost"]);
                    PriceRatio = Convert.ToDecimal(PRows[0]["PriceRatio"]);
                    DigitCapacity = Convert.ToInt32(PRows[0]["DigitCapacity"]);
                    OriginalPrice = MarketingCost * PriceRatio * (1 + PriceGroup);
                    OriginalPrice = Decimal.Round(OriginalPrice, DigitCapacity, MidpointRounding.AwayFromZero);
                    CurrentPrice = DiscountVolume * OriginalPrice;

                    if (ClientID == 145)
                    {
                        OriginalPrice = Convert.ToDecimal(PRows[0]["ZOVPrice"]);
                        CurrentPrice = DiscountVolume * OriginalPrice;
                    }
                    Price = CurrentPrice - CurrentPrice / 100 * TotalDiscount;
                    if (Sale)
                    {
                        decimal SampleSaleValue = GetSampleDiscount(Convert.ToInt32(DecorOrderRow["DecorConfigID"]));
                        decimal SaleValue = GetDiscount(Convert.ToInt32(DecorOrderRow["DecorConfigID"]));
                        if (IsSample)
                            Price = (CurrentPrice - CurrentPrice / 100 * (50 + SampleSaleValue)) * (1 - DiscountDirector / 100);
                        else
                            Price = (CurrentPrice - CurrentPrice / 100 * (TotalDiscount + SaleValue)) * (1 - DiscountDirector / 100);
                    }
                    else
                    {
                        if (IsSample)
                            Price = (CurrentPrice - CurrentPrice / 100 * 50) * (1 - DiscountDirector / 100);
                        else
                            Price = (CurrentPrice - CurrentPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    }
                    //if (IsSample)
                    //{
                    //    Price = (CurrentPrice - CurrentPrice * 50 / 100) * (1 - DiscountDirector / 100);
                    //    //Price = OriginalPrice - OriginalPrice / 100 * (50 + DiscountDirector);
                    //}
                    //else
                    //{
                    //    Price = (CurrentPrice - CurrentPrice / 100 * TotalDiscount) * (1 - DiscountDirector / 100);
                    //    //Price = CurrentPrice - CurrentPrice / 100 * (TotalDiscount + DiscountDirector);
                    //}
                    Price = Decimal.Round(Price, 2, MidpointRounding.AwayFromZero);
                }
            }
        }
        
        private int GetMeasureType(DataRow DecorOrderRow)
        {
            DataRow[] Rows = DecorConfigDataTable.Select("DecorConfigID = " + DecorOrderRow["DecorConfigID"].ToString());

            return Convert.ToInt32(Rows[0]["MeasureID"]);
        }

        private decimal GetTotalVolume(string TechStoreName, bool BaseProfile)
        {
            decimal Height = -1;
            decimal Length = -1;
            decimal Count = -1;
            decimal TotalVolume = 0;
            string s1 = string.Empty;
            string s2 = string.Empty;
            string s3 = string.Empty;
            string s4 = string.Empty;
            int index1 = TechStoreName.IndexOf("-");
            int index2 = TechStoreName.IndexOfAny(new char[] { '-' }, index1 + 1);
            int index3 = TechStoreName.IndexOfAny(new char[] { ' ' }, index1 + 1);
            int index4 = TechStoreName.IndexOfAny(new char[] { ' ' }, index3 + 1);
            if (index4 == -1)
                index4 = TechStoreName.Length;
            if (index1 != -1)
            {
                if (index2 != -1)
                    s1 = TechStoreName.Substring(index1 + 1, index2 - index1 - 1);
                else
                {
                    if (index3 != -1)
                        s1 = TechStoreName.Substring(index1 + 1, index3 - index1 - 1);
                }
                if (index3 != -1 && index4 != -1)
                    s2 = TechStoreName.Substring(index3 + 1, index4 - index3 - 1);
            }

            if (BaseProfile)
            {
                for (int i = 0; i < AllDecorOrdersTable.Rows.Count; i++)
                {
                    Height = Convert.ToDecimal(AllDecorOrdersTable.Rows[i]["Height"]);
                    Length = Convert.ToDecimal(AllDecorOrdersTable.Rows[i]["Length"]);
                    Count = Convert.ToInt32(AllDecorOrdersTable.Rows[i]["Count"]);
                    decimal MinBalanceOnStorage = 0;
                    string Name = string.Empty;

                    Tuple<string, decimal> tuple = TablesManager.GetTechStoreNameAndBalance(
                        Convert.ToInt32(AllDecorOrdersTable.Rows[i]["DecorConfigID"]));
                    Name = tuple.Item1;
                    MinBalanceOnStorage = tuple.Item2;
                    //GetTechStoreNameAndBalance(Convert.ToInt32(AllDecorOrdersTable.Rows[i]["DecorConfigID"]), ref Name, ref MinBalanceOnStorage);
                    if (Name.Length > 0)
                    {
                        if (MinBalanceOnStorage != 0)
                        {
                            index1 = Name.IndexOf("-");
                            index2 = Name.IndexOfAny(new char[] { '-' }, index1 + 1);
                            index3 = Name.IndexOfAny(new char[] { ' ' }, index1 + 1);
                            index4 = Name.IndexOfAny(new char[] { ' ' }, index3 + 1);
                            if (index4 == -1)
                                index4 = Name.Length;
                            if (index1 != -1)
                            {
                                if (index2 != -1)
                                    s3 = Name.Substring(index1 + 1, index2 - index1 - 1);
                                else
                                {
                                    if (index3 != -1)
                                        s3 = Name.Substring(index1 + 1, index3 - index1 - 1);
                                }
                                if (index3 != -1 && index4 != -1)
                                    s4 = Name.Substring(index3 + 1, index4 - index3 - 1);
                                if (s1 == s3 && s2 == s4)
                                {
                                    if (Height != -1)
                                        TotalVolume += Height * Count / 1000;
                                    if (Length != -1)
                                        TotalVolume += Length * Count / 1000;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < AllDecorOrdersTable.Rows.Count; i++)
                {
                    Height = Convert.ToDecimal(AllDecorOrdersTable.Rows[i]["Height"]);
                    Length = Convert.ToDecimal(AllDecorOrdersTable.Rows[i]["Length"]);
                    Count = Convert.ToInt32(AllDecorOrdersTable.Rows[i]["Count"]);
                    decimal MinBalanceOnStorage = 0;
                    string Name = string.Empty;

                    Tuple<string, decimal> tuple = TablesManager.GetTechStoreNameAndBalance(
                        Convert.ToInt32(AllDecorOrdersTable.Rows[i]["DecorConfigID"]));
                    Name = tuple.Item1;
                    MinBalanceOnStorage = tuple.Item2;

                    //GetTechStoreNameAndBalance(Convert.ToInt32(AllDecorOrdersTable.Rows[i]["DecorConfigID"]), ref Name, ref MinBalanceOnStorage);
                    if (Name.Length > 0)
                    {
                        if (MinBalanceOnStorage == 0)
                        {
                            index1 = Name.IndexOf("-");
                            index2 = Name.IndexOfAny(new char[] { '-' }, index1 + 1);
                            index3 = Name.IndexOfAny(new char[] { ' ' }, index1 + 1);
                            index4 = Name.IndexOfAny(new char[] { ' ' }, index3 + 1);
                            if (index4 == -1)
                                index4 = Name.Length;
                            if (index1 != -1)
                            {
                                if (index2 != -1)
                                    s3 = Name.Substring(index1 + 1, index2 - index1 - 1);
                                else
                                {
                                    if (index3 != -1)
                                        s3 = Name.Substring(index1 + 1, index3 - index1 - 1);
                                }
                                if (index3 != -1 && index4 != -1)
                                    s4 = Name.Substring(index3 + 1, index4 - index3 - 1);
                                if (s1 == s3 && s2 == s4)
                                {
                                    if (Height != -1)
                                        TotalVolume += Height * Count / 1000;
                                    if (Length != -1)
                                        TotalVolume += Length * Count / 1000;
                                }
                            }
                        }
                    }
                }
            }

            return TotalVolume;
        }

        private void GetTechStoreNameAndBalance(int DecorConfigID, ref string TechStoreName, ref decimal MinBalanceOnStorage)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT DecorConfig.MinBalanceOnStorage, Decor.TechStoreName FROM DecorConfig INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON DecorConfig.DecorID = Decor.TechStoreID WHERE DecorConfigID=" + DecorConfigID,
                ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        TechStoreName = DT.Rows[0]["TechStoreName"].ToString();
                        MinBalanceOnStorage = Convert.ToInt32(DT.Rows[0]["MinBalanceOnStorage"]);
                    }
                }

            }
        }

        private string GetTechStoreName(int DecorID)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TechStoreName FROM TechStore WHERE TechStoreID=" + DecorID,
                ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return DT.Rows[0]["TechStoreName"].ToString();
                }

            }
            return string.Empty;
        }

        private void ProfilCalculateItem(int clientId, DataRow decorOrderRow, decimal DiscountDirector, decimal TotalDiscount,
            int CurrencyTypeID, decimal currency,
            ref decimal originalCost, ref decimal cost, ref decimal currencyOrderCost)
        {
            decimal OriginalDecorPrice = 0;
            decimal price = 0;
            int MeasureType = -1;

            decimal Height = -1;
            decimal Width = -1;
            decimal length = -1;
            decimal Count = -1;

            bool BaseProfile = false;
            decimal DiscountVolume = Convert.ToDecimal(decorOrderRow["DiscountVolume"]);
            decimal MinBalanceOnStorage = 0;
            string TechStoreName = string.Empty;
            int DecorConfigID = Convert.ToInt32(decorOrderRow["DecorConfigID"]);
            DataRow[] Rows = DecorConfigDataTable.Select("DecorConfigID = " + decorOrderRow["DecorConfigID"].ToString());
            if (Rows.Count() > 0)
            {
                MinBalanceOnStorage = Convert.ToDecimal(Rows[0]["MinBalanceOnStorage"]);
                if (MinBalanceOnStorage != 0)
                    BaseProfile = true;
                TechStoreName = TablesManager.GetTechStoreNameByDecorId(Convert.ToInt32(decorOrderRow["DecorID"]));

                //TechStoreName = GetTechStoreName(Convert.ToInt32(DecorOrderRow["DecorID"]));
            }
            MeasureType = GetMeasureType(decorOrderRow);

            try
            {
                Height = Convert.ToDecimal(decorOrderRow["Height"]);
            }
            catch
            { }

            try
            {
                length = Convert.ToDecimal(decorOrderRow["Length"]);
            }
            catch { }

            try
            {
                Width = Convert.ToDecimal(decorOrderRow["Width"]);
            }
            catch { }

            try
            {
                Count = Convert.ToDecimal(decorOrderRow["Count"]);
            }
            catch
            {
            }

            if (MeasureType == 1)
            {
                if (NeedCalcPrice)
                    GetProfilDecorPrice(clientId, decorOrderRow, DiscountDirector, TotalDiscount, ref OriginalDecorPrice, ref price);
                else
                {
                    OriginalDecorPrice = Convert.ToDecimal(decorOrderRow["OriginalPrice"]);
                    price = Convert.ToDecimal(decorOrderRow["Price"]);
                }
                if (Height != -1)
                {             //м.кв.
                    if (CurrencyTypeID == 0)
                    {
                        decimal d = price * currency;
                        d = Math.Ceiling(d / 100.0m) * 100.0m;
                        currencyOrderCost = Decimal.Round(Height * Width / 1000000, 3, MidpointRounding.AwayFromZero) * Count * d;
                    }
                    else
                    {
                        decimal d = price * currency;
                        currencyOrderCost = Decimal.Round(Height * Width / 1000000, 3, MidpointRounding.AwayFromZero) * Count * d;
                    }
                    originalCost = Decimal.Round(Height * Width / 1000000, 3, MidpointRounding.AwayFromZero) * Count * OriginalDecorPrice;
                    cost = Decimal.Round(Height * Width / 1000000, 3, MidpointRounding.AwayFromZero) * Count * price;
                }
                if (length != -1)
                {
                    if (CurrencyTypeID == 0)
                    {
                        decimal d = price * currency;
                        d = Math.Ceiling(d / 100.0m) * 100.0m;
                        currencyOrderCost = Decimal.Round(length * Width / 1000000, 3, MidpointRounding.AwayFromZero) * Count * d;
                    }
                    else
                    {
                        decimal d = price * currency;
                        currencyOrderCost = Decimal.Round(length * Width / 1000000, 3, MidpointRounding.AwayFromZero) * Count * d;
                    }
                    originalCost = Decimal.Round(length * Width / 1000000, 3, MidpointRounding.AwayFromZero) * Count * OriginalDecorPrice;
                    cost = Decimal.Round(length * Width / 1000000, 3, MidpointRounding.AwayFromZero) * Count * price;
                }
            }

            if (MeasureType == 2)
            {
                decimal TotalVolume = 0;
                if (MinBalanceOnStorage == 0)
                {
                    TotalVolume = GetTotalVolume(TechStoreName, BaseProfile);
                    DiscountVolume = GetDiscountVolume(TotalVolume);
                    if (decorOrderRow["DecorID"].ToString() == "15523")//плк-110
                        GetPilastrPrice(clientId, decorOrderRow, length, DiscountDirector, TotalDiscount, ref OriginalDecorPrice, ref price, DiscountVolume);
                    else
                        GetProfilDecorPrice(clientId, decorOrderRow, DiscountDirector, TotalDiscount, ref OriginalDecorPrice, ref price, DiscountVolume);
                    decorOrderRow["DiscountVolume"] = DiscountVolume;
                }
                else
                {
                    TotalVolume = GetTotalVolume(TechStoreName, BaseProfile);
                    if (TotalVolume >= 2400)
                    {
                        DiscountVolume = GetDiscountVolume(TotalVolume);
                        if (decorOrderRow["DecorID"].ToString() == "15523")//плк-110
                            GetPilastrPrice(clientId, decorOrderRow, length, DiscountDirector, TotalDiscount, ref OriginalDecorPrice, ref price, DiscountVolume);
                        else
                            GetProfilDecorPrice(clientId, decorOrderRow, DiscountDirector, TotalDiscount, ref OriginalDecorPrice, ref price, DiscountVolume);
                        decorOrderRow["DiscountVolume"] = DiscountVolume;
                    }
                    else
                    {
                        decorOrderRow["DiscountVolume"] = 1;
                        if (decorOrderRow["DecorID"].ToString() == "15523")//плк-110
                            GetPilastrPrice(clientId, decorOrderRow, length, DiscountDirector, TotalDiscount, ref OriginalDecorPrice, ref price);
                        else
                            GetProfilDecorPrice(clientId, decorOrderRow, DiscountDirector, TotalDiscount, ref OriginalDecorPrice, ref price);
                    }
                }
                if (Height != -1)
                {
                    if (CurrencyTypeID == 0)
                    {
                        decimal d = price * currency;
                        //d = Math.Ceiling(d / 100.0m) * 100.0m;
                        currencyOrderCost = Decimal.Round(Height / 1000, 3, MidpointRounding.AwayFromZero) * Count * d;
                    }
                    else
                    {
                        decimal d = price * currency;
                        currencyOrderCost = Decimal.Round(Height / 1000, 3, MidpointRounding.AwayFromZero) * Count * d;
                    }
                    originalCost = Decimal.Round(Height / 1000, 3, MidpointRounding.AwayFromZero) * Count * OriginalDecorPrice;
                    cost = Decimal.Round(Height / 1000, 3, MidpointRounding.AwayFromZero) * Count * price;
                }
                if (length != -1)
                {
                    if (CurrencyTypeID == 0)
                    {
                        decimal d = price * currency;
                        //d = Math.Ceiling(d / 100.0m) * 100.0m;
                        currencyOrderCost = Decimal.Round(length / 1000, 3, MidpointRounding.AwayFromZero) * Count * d;
                    }
                    else
                    {
                        decimal d = price * currency;
                        currencyOrderCost = Decimal.Round(length / 1000, 3, MidpointRounding.AwayFromZero) * Count * d;
                    }
                    originalCost = Decimal.Round(length / 1000, 3, MidpointRounding.AwayFromZero) * Count * OriginalDecorPrice;
                    cost = Decimal.Round(length / 1000, 3, MidpointRounding.AwayFromZero) * Count * price;
                }
            }

            if (MeasureType == 3)
            {
                if (NeedCalcPrice)
                    GetProfilDecorPrice(clientId, decorOrderRow, DiscountDirector, TotalDiscount, ref OriginalDecorPrice, ref price);            //шт.
                else
                {
                    OriginalDecorPrice = Convert.ToDecimal(decorOrderRow["OriginalPrice"]);
                    price = Convert.ToDecimal(decorOrderRow["Price"]);
                }
                if (CurrencyTypeID == 0)
                {
                    decimal d = price * currency;
                    d = Math.Ceiling(d / 100.0m) * 100.0m;
                    currencyOrderCost = Count * d;
                }
                else
                {
                    decimal d = price * currency;
                    currencyOrderCost = Count * d;
                }
                originalCost = Count * OriginalDecorPrice;
                cost = Count * price;
            }

            originalCost = Math.Ceiling(originalCost / 0.01m) * 0.01m;
            cost = Math.Ceiling(cost / 0.01m) * 0.01m;
            if (NeedCalcPrice)
            {
                decorOrderRow["ClientOriginalPrice"] = Math.Ceiling(OriginalDecorPrice * currency / 0.01m) * 0.01m;
                decorOrderRow["OriginalPrice"] = OriginalDecorPrice;
                decorOrderRow["Price"] = price;
            }
            bool isSample = Convert.ToBoolean(decorOrderRow["IsSample"]);
            //if (IsSample)
            //{
            //    if (OriginalDecorPrice != 0)
            //        DecorOrderRow["TotalDiscount"] = (1 - Price / OriginalDecorPrice) * 100;
            //}
            //else
            //    DecorOrderRow["TotalDiscount"] = TotalDiscount + DiscountDirector;
            if (OriginalDecorPrice != 0)
            {
                decimal d = Decimal.Round((1 - price / OriginalDecorPrice) * 100, 2, MidpointRounding.AwayFromZero);
                decorOrderRow["TotalDiscount"] = d;
            }

            if (decorOrderRow["DecorID"].ToString() == "15523")//плк-110
            {
                decimal PilastrPrice = 1;
                //GetPilastrPrice(ClientID, DecorOrderRow, Length, DiscountDirector, TotalDiscount, ref PilastrPrice);
                if (clientId == 145)
                {
                    if (isSample)
                    {
                        currencyOrderCost =
                            decimal.Round(
                                Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price * currency), 3,
                                MidpointRounding.AwayFromZero);
                        originalCost =
                            decimal.Round(
                                Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price), 3,
                                MidpointRounding.AwayFromZero);
                        cost = decimal.Round(Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price), 3,
                            MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        currencyOrderCost =
                            decimal.Round(
                                Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price * currency), 3,
                                MidpointRounding.AwayFromZero);
                        originalCost =
                            decimal.Round(
                                Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price), 3,
                                MidpointRounding.AwayFromZero);
                        cost = decimal.Round(Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price), 3,
                            MidpointRounding.AwayFromZero);
                    }
                }
                else
                {
                    currencyOrderCost =
                        decimal.Round(Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price * currency), 3,
                            MidpointRounding.AwayFromZero);
                    originalCost =
                        decimal.Round(Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price), 3,
                            MidpointRounding.AwayFromZero);
                    cost = decimal.Round(Convert.ToInt32(decorOrderRow["Count"]) * (length / 1000 * price), 3,
                        MidpointRounding.AwayFromZero);
                }
            }

            decorOrderRow["OriginalCost"] = originalCost;
            decorOrderRow["Cost"] = cost;
            decorOrderRow["CurrencyTypeID"] = CurrencyTypeID;
            currencyOrderCost = Math.Ceiling(currencyOrderCost / 0.01m) * 0.01m;
            //CurrencyOrderCost = Decimal.Round(CurrencyOrderCost, 2, MidpointRounding.AwayFromZero);
            //if (CurrencyTypeID == 4)
            //{
            //    CurrencyOrderCost = Math.Ceiling(CurrencyOrderCost / 100m) * 100m;
            //}
            decorOrderRow["CurrencyCost"] = currencyOrderCost;
        }
        
        public decimal GetDecorWeight(DataRow DecorOrderRow)
        {
            if (DecorOrderRow["Weight"] == DBNull.Value)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка конфигурации: нет веса в DecorConfigID=" + DecorOrderRow["DecorConfigID"].ToString() +
                    ". Вес будет выставлен в 0.");
                return 0;
            }
            decimal Weight = 0;

            DataRow[] Row = DecorConfigDataTable.Select("DecorConfigID = " + DecorOrderRow["DecorConfigID"].ToString());

            if (Row[0]["Weight"] == DBNull.Value)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка конфигурации: нет веса в DecorConfigID=" + DecorOrderRow["DecorConfigID"].ToString() +
                    ". Вес будет выставлен в 0.");
                return 0;
            }
            if (Row[0]["Weight"] != DBNull.Value)
            {
                if (Row[0]["WeightMeasureID"].ToString() == "1")
                {
                    if (Convert.ToInt32(DecorOrderRow["Height"]) != -1)
                        Weight = Convert.ToDecimal(DecorOrderRow["Height"]) * Convert.ToDecimal(DecorOrderRow["Width"]) / 1000000
                             * Convert.ToDecimal(Row[0]["Weight"]) * Convert.ToDecimal(DecorOrderRow["Count"]);
                    if (Convert.ToInt32(DecorOrderRow["Length"]) != -1)
                        Weight = Convert.ToDecimal(DecorOrderRow["Length"]) * Convert.ToDecimal(DecorOrderRow["Width"]) / 1000000
                            * Convert.ToDecimal(Row[0]["Weight"]) * Convert.ToDecimal(DecorOrderRow["Count"]);
                }
                if (Row[0]["WeightMeasureID"].ToString() == "2")
                {
                    if (Convert.ToInt32(DecorOrderRow["Height"]) != -1)
                        Weight = Convert.ToDecimal(DecorOrderRow["Height"]) / 1000 * Convert.ToDecimal(Row[0]["Weight"]) *
                                 Convert.ToDecimal(DecorOrderRow["Count"]);
                    if (Convert.ToInt32(DecorOrderRow["Length"]) != -1)
                        Weight = Convert.ToDecimal(DecorOrderRow["Length"]) / 1000 * Convert.ToDecimal(Row[0]["Weight"]) *
                             Convert.ToDecimal(DecorOrderRow["Count"]);
                }
                if (Row[0]["WeightMeasureID"].ToString() == "3")
                    Weight = Convert.ToDecimal(Row[0]["Weight"]) * Convert.ToDecimal(DecorOrderRow["Count"]);
            }
            Weight = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
            DecorOrderRow["ItemWeight"] = Weight / Convert.ToDecimal(DecorOrderRow["Count"]);
            DecorOrderRow["Weight"] = Weight;

            return Weight;
        }

        public void GetAllDecorOrders(int MegaOrderID)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorOrders WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID = " + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    AllDecorOrdersTable.Clear();
                    DA.Fill(AllDecorOrdersTable);
                }
            }
        }

        public void CalculateDecor(int ClientID, int MainOrderID, decimal ProfilDiscountDirector, decimal TPSDiscountDirector,
            decimal ProfilTotalDiscount, decimal TPSTotalDiscount, int CurrencyTypeID, decimal Currency,
            ref decimal OriginalProfilCost, ref decimal OriginalTPSCost, ref decimal OriginalTotalCost,
            ref decimal ProfilCost, ref decimal TPSCost, ref decimal TotalCost, ref decimal TotalCurrencyCost,
            ref decimal DecorOrderWeight, object ConfirmDateTime)
        {
            GetStocks(ClientID, ConfirmDateTime);
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorOrders WHERE MainOrderID = " + MainOrderID,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    DecorOrdersTable.Clear();
                    DA.Fill(DecorOrdersTable);
                    //DT.TableName = DecorCatalogOrder.DecorProductsDataTable.Rows[i]["OrdersTableName"].ToString();

                    if (DecorOrdersTable.Rows.Count == 0)
                        return;

                    foreach (DataRow Row in DecorOrdersTable.Rows)
                    {
                        int FactoryID = Convert.ToInt32(Row["FactoryID"]);
                        decimal OriginalCost = 0;
                        decimal Cost = 0;
                        decimal CurrencyCost = 0;

                        DecorOrderWeight += GetDecorWeight(Row);

                        if (FactoryID == 1)
                        {
                            ProfilCalculateItem(ClientID, Row, ProfilDiscountDirector, ProfilTotalDiscount, CurrencyTypeID, Currency, ref OriginalCost, ref Cost, ref CurrencyCost);
                            OriginalProfilCost += OriginalCost;
                            ProfilCost += Cost;
                        }
                        if (FactoryID == 2)
                        {
                            //TPSCalculateItem(ClientID, Row, TPSDiscountDirector, TPSTotalDiscount, CurrencyTypeID, Currency, ref OriginalCost, ref Cost, ref CurrencyCost);
                            //OriginalTPSCost += OriginalCost;
                            //TPSCost += Cost;
                        }
                        TotalCurrencyCost += CurrencyCost;
                    }
                    OriginalTotalCost = OriginalProfilCost + OriginalTPSCost;
                    TotalCost = ProfilCost + TPSCost;
                    DA.Update(DecorOrdersTable);
                }
            }
        }

        public void AssignTransportCost(int MegaOrderID, decimal ProfilComplaintCost, decimal TPSComplaintCost, decimal TransportAdditionalCost)
        {
            decimal ProfilTotalWeight = 0;
            decimal TPSTotalWeight = 0;
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow Row in DT.Rows)
                            {
                                decimal Square = Convert.ToDecimal(Row["Square"]);
                                decimal Weight = Convert.ToDecimal(Row["Weight"]);
                                if (Square > 0)
                                    Weight = Weight + Square * Convert.ToDecimal(0.7);
                                Weight = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
                                ProfilTotalWeight += Weight;
                            }
                        }
                    }
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorOrders WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow Row in DT.Rows)
                                ProfilTotalWeight += Convert.ToDecimal(Row["Weight"]);
                        }
                    }
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow Row in DT.Rows)
                            {
                                decimal Square = Convert.ToDecimal(Row["Square"]);
                                decimal Weight = Convert.ToDecimal(Row["Weight"]);
                                if (Square > 0)
                                    Weight = Weight + Square * Convert.ToDecimal(0.7);
                                Weight = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
                                TPSTotalWeight += Weight;
                            }
                        }
                    }
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorOrders WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow Row in DT.Rows)
                                TPSTotalWeight += Convert.ToDecimal(Row["Weight"]);
                        }
                    }
                }
            }
            decimal ProfilTransportAdditionalCost = 0;
            decimal TPSTransportAdditionalCost = 0;
            if (ProfilTotalWeight > 0)
                ProfilTransportAdditionalCost = ProfilComplaintCost + TransportAdditionalCost * ProfilTotalWeight / (ProfilTotalWeight + TPSTotalWeight);
            if (TPSTotalWeight > 0)
                TPSTransportAdditionalCost = TPSComplaintCost + TransportAdditionalCost * TPSTotalWeight / (ProfilTotalWeight + TPSTotalWeight);
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorOrders WHERE MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID=" + MegaOrderID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DecorOrdersTable = new DataTable())
                    {
                        DA.Fill(DecorOrdersTable);

                        if (DecorOrdersTable.Rows.Count == 0)
                            return;

                        foreach (DataRow Row in DecorOrdersTable.Rows)
                        {
                            decimal PriceWithTransport = 0;
                            decimal CostWithTransport = 0;
                            int FactoryID = Convert.ToInt32(Row["FactoryID"]);
                            decimal Price = Convert.ToDecimal(Row["Price"]);
                            decimal Cost = Convert.ToDecimal(Row["Cost"]);
                            decimal Weight = Convert.ToDecimal(Row["Weight"]);

                            decimal d = 0;
                            decimal d1 = 0;

                            if (FactoryID == 1)
                            {
                                if (ProfilTotalWeight == 0)
                                    continue;
                                d = Weight / (ProfilTotalWeight / 100);
                                d1 = ProfilTransportAdditionalCost / 100 * d;
                            }
                            if (FactoryID == 2)
                            {
                                if (TPSTotalWeight == 0)
                                    continue;
                                d = Weight / (TPSTotalWeight / 100);
                                d1 = TPSTransportAdditionalCost / 100 * d;
                            }
                            DataRow[] row = DecorConfigDataTable.Select("DecorConfigID = " + Row["DecorConfigID"].ToString());
                            if (row[0]["Weight"] != DBNull.Value)
                            {
                                decimal d2 = 0;
                                if (row[0]["WeightMeasureID"].ToString() == "1")
                                {
                                    if (Convert.ToInt32(Row["Height"]) != -1)
                                        d2 = Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Width"]) / 1000000
                                             * Convert.ToDecimal(Row["Count"]);
                                    if (Convert.ToInt32(Row["Length"]) != -1)
                                        d2 = Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Width"]) / 1000000
                                            * Convert.ToDecimal(Row["Count"]);
                                }
                                if (row[0]["WeightMeasureID"].ToString() == "2")
                                {
                                    if (Convert.ToInt32(Row["Height"]) != -1)
                                        d2 = Convert.ToDecimal(Row["Height"]) / 1000 *
                                                 Convert.ToDecimal(Row["Count"]);
                                    if (Convert.ToInt32(Row["Length"]) != -1)
                                        d2 = Convert.ToDecimal(Row["Length"]) / 1000 *
                                             Convert.ToDecimal(Row["Count"]);
                                }
                                if (row[0]["WeightMeasureID"].ToString() == "3")
                                {

                                    d2 = Convert.ToDecimal(Row["Count"]);
                                }
                                if (d2 == 0)
                                {
                                    System.Windows.Forms.MessageBox.Show("Кол-во продукции не может быть равно 0");
                                }
                                else
                                {
                                    PriceWithTransport = Price + d1 / d2;
                                }
                            }

                            PriceWithTransport = Decimal.Round(PriceWithTransport, 2, MidpointRounding.AwayFromZero);
                            Row["PriceWithTransport"] = PriceWithTransport;
                            CostWithTransport = Cost + d1;
                            CostWithTransport = Decimal.Round(CostWithTransport, 2, MidpointRounding.AwayFromZero);
                            Row["CostWithTransport"] = CostWithTransport;
                        }

                        DA.Update(DecorOrdersTable);
                    }
                }
            }
        }


        private void GetStocks(int ClientID, object ConfirmDateTime)
        {
            Sale = false;
            DateTime DocDateTime = DateTime.Now;
            if (ConfirmDateTime != DBNull.Value)
                DocDateTime = Convert.ToDateTime(ConfirmDateTime);

            if (StockDetailsDT == null)
                StockDetailsDT = new DataTable();
            else
                StockDetailsDT.Clear();

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT StockDetails.*, Stocks.Discount, Stocks.SampleDiscount FROM StockDetails INNER JOIN Stocks ON StockDetails.StockID=Stocks.StockID AND Stocks.Enable=1 AND 
                CAST(Stocks.FirstDate AS Date) <= '" + DocDateTime.ToString("yyyy-MM-dd") +
               "' AND CAST(Stocks.LastDate AS Date) >= '" + DocDateTime.ToString("yyyy-MM-dd") + "' WHERE StockDetails.StockID IN (SELECT StockID FROM StockClients WHERE ClientID=" + ClientID + ") AND StockDetails.ProductType=1",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                if (DA.Fill(StockDetailsDT) > 0)
                    Sale = true;
            }
        }

        private decimal GetDiscount(int ConfigID)
        {
            DataRow[] rows = StockDetailsDT.Select("ConfigID = " + ConfigID);
            if (rows.Count() > 0)
                return Convert.ToDecimal(rows[0]["Discount"]);
            return 0;
        }

        private decimal GetSampleDiscount(int ConfigID)
        {
            DataRow[] rows = StockDetailsDT.Select("ConfigID = " + ConfigID);
            if (rows.Count() > 0)
                return Convert.ToDecimal(rows[0]["SampleDiscount"]);
            return 0;
        }
    }
}
