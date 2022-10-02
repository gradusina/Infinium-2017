
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace Infinium.Modules.ProfileAssignments.Planning
{
    public class ClientsManager
    {
        public BindingSource clientsBs;
        public BindingSource orderNumbersBs;
        private readonly DataTable clientsDt;
        private readonly DataTable orderNumbersDt;

        public ClientsManager()
        {
            clientsDt = new DataTable();
            orderNumbersDt = new DataTable();

            clientsBs = new BindingSource
            {
                DataSource = clientsDt
            };
            orderNumbersBs = new BindingSource
            {
                DataSource = orderNumbersDt
            };

            var selectCommand = @"SELECT ClientID, ClientName FROM Clients WHERE Enabled=1 ORDER BY ClientName";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingReferenceConnectionString))
            {
                da.Fill(clientsDt);
            }

            selectCommand = @"SELECT TOP 0 ClientID, MegaOrderID, OrderNumber FROM MegaOrders";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                da.Fill(orderNumbersDt);
            }
        }

        public void FillOrderNumberDt(int clientId)
        {
            orderNumbersDt.Clear();

            var selectCommand =
                $@"SELECT ClientID, MegaOrderID, OrderNumber FROM MegaOrders where clientId={clientId} ORDER BY OrderNumber DESC";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                da.Fill(orderNumbersDt);
            }
        }

        public void FilterOrderNumbers(int clientId)
        {
            orderNumbersBs.Filter = "clientId=" + clientId;
        }

        public string GetClientName(int id)
        {
            var name = "no_name";

            var rows = clientsDt.Select("ClientID=" + id);
            if (rows.Any())
                name = rows[0]["ClientName"].ToString();
            return name;
        }
    }

    public class MoldingsSectionPlanning
    {
        private const decimal M3WidthConst = 6.5m;

        private readonly DataSet OrdersDs;
        private readonly DataTable _decorAssignments;
        private readonly DataTable _dbOrdersDt;
        private readonly DataTable _dbProdOrdersDt;
        private readonly DataTable _displayOrdersDt;
        private readonly DataTable _displayProdOrdersDt;
        private readonly DataTable _moldingsSectionToProduceDt;
        private DataTable _colorsDt;
        private DataTable _decorDt;
        private DataTable _usersDt;
        private DataTable _millingMachinesDt;
        private DataTable _facingMachinesDt;
        private BindingSource _millingMachinesBs;
        private BindingSource _facingMachinesBs;

        public MoldingsSectionPlanning()
        {
            OrdersDs = new DataSet();
            _decorAssignments = new DataTable();
            _dbOrdersDt = new DataTable();
            _dbProdOrdersDt = new DataTable();
            _displayOrdersDt = new DataTable();
            _displayProdOrdersDt = new DataTable();
            _moldingsSectionToProduceDt = new DataTable();

            var fi = typeof(OrderRow).GetFields(BindingFlags.Public | BindingFlags.Instance);
            foreach (var info in fi) _displayOrdersDt.Columns.Add(new DataColumn(info.Name, info.FieldType));

            fi = typeof(ProdOrderRow).GetFields(BindingFlags.Public | BindingFlags.Instance);
            foreach (var info in fi) _displayProdOrdersDt.Columns.Add(new DataColumn(info.Name, info.FieldType));

            GetDecorDt();
            GetColorsDt();
            GetUsersDt();
            GetMillingMachinesDt();
            GetFacingMachinesDt();

            _millingMachinesBs = new BindingSource()
            {
                DataSource = _millingMachinesDt
            };
            _facingMachinesBs = new BindingSource()
            {
                DataSource = _facingMachinesDt
            };
        }

        public int DistinctThicknessTables => OrdersDs.Tables.Count;

        public void ClearOrders()
        {
            _displayOrdersDt.Clear();
        }

        public void ClearProdOrders()
        {
            _displayProdOrdersDt.Clear();
        }

        public DataTable CurrentTable(int index)
        {
            return OrdersDs.Tables[index];
        }

        public DataTable CurrentProdOrdersTable()
        {
            return _displayProdOrdersDt;
        }

        public int CurrentTableThickness(int index)
        {
            var thickness = 0;
            if (OrdersDs.Tables[index].Rows[0]["Thickness"] != DBNull.Value)
                thickness = Convert.ToInt32(OrdersDs.Tables[index].Rows[0]["Thickness"]);
            return thickness;
        }

        public void SaveProdOrdersToDb()
        {
            DataTable changesDt = _displayProdOrdersDt.GetChanges();


            var selectCommand = @"SELECT * FROM MoldingsSectionToProduce";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dt = new DataTable())
                    {
                        if (da.Fill(dt) <= 0) return;

                        for (var i = 0; i < dt.Rows.Count; i++)
                        {
                            int id = Convert.ToInt32(dt.Rows[i]["id"]);
                            int priority = 0;
                            int millingOption = 0;
                            int facingOption = 0;

                            var rows = _displayProdOrdersDt.Select("id=" + id);
                            if (rows.Any())
                            {
                                if (rows[0]["priority"] != DBNull.Value)
                                    priority = Convert.ToInt32(rows[0]["priority"]);
                                if (rows[0]["millingOption"] != DBNull.Value)
                                    millingOption = Convert.ToInt32(rows[0]["millingOption"]);
                                if (rows[0]["facingOption"] != DBNull.Value)
                                    facingOption = Convert.ToInt32(rows[0]["facingOption"]);
                            }

                            dt.Rows[i]["priority"] = priority;
                            dt.Rows[i]["millingOption"] = millingOption;
                            dt.Rows[i]["facingOption"] = facingOption;
                        }

                        da.Update(dt);
                    }
                }
            }
        }

        public void SaveOrdersToDb()
        {
            var selectCommand = @"SELECT * FROM MoldingsSectionToProduce";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dt = new DataTable())
                    {
                        if (da.Fill(dt) <= 0) return;

                        int maxPriority = Convert.ToInt32(dt.Compute("MAX(Priority)", "")) + 1;

                        for (int i = 0; i < OrdersDs.Tables.Count; i++)
                        {
                            for (int j = 0; j < OrdersDs.Tables[i].Rows.Count; j++)
                            {
                                if (!Convert.ToBoolean(OrdersDs.Tables[i].Rows[j]["ToProduce"]))
                                {
                                    var rows1 = dt.Select("DecorOrderID=" +
                                                          Convert.ToInt32(OrdersDs.Tables[i].Rows[j]["DecorOrderID"]));
                                    if (rows1.Any())
                                        rows1[0].Delete();
                                    continue;
                                }

                                var rows = dt.Select("DecorOrderID=" +
                                                     Convert.ToInt32(OrdersDs.Tables[i].Rows[j]["DecorOrderID"]));
                                if (rows.Any())
                                    continue;

                                DataRow newRow = dt.NewRow();
                                newRow["DecorOrderID"] = Convert.ToInt32(OrdersDs.Tables[i].Rows[j]["DecorOrderID"]);
                                newRow["ProduceCount"] = Convert.ToInt32(OrdersDs.Tables[i].Rows[j]["ProduceCount"]);
                                newRow["Priority"] = maxPriority++;
                                dt.Rows.Add(newRow);
                            }
                        }

                        da.Update(dt);
                    }
                }
            }
        }

        public int TotalProduceCount(int index)
        {
            var count = 0;
            for (var i = 0; i < OrdersDs.Tables[index].Rows.Count; i++)
                count += Convert.ToInt32(OrdersDs.Tables[index].Rows[i]["ProduceCount"]);

            return count;
        }

        public void UpdateOrders(ClientsManager clientsManager, bool groupByThickness, bool onProd, bool inProd, bool onApproval, bool allClients, bool showZero, int clientId)
        {
            GetSavedOrdersFromDb();
            GetOrdersFromDb(onProd, inProd, onApproval, allClients, clientId);
            FillOrderRows(clientsManager, showZero);
            if (groupByThickness)
                GroupByThickness();
            else
                UnGroup();
        }

        public void UpdateProdOrders()
        {
            GetProdOrdersFromDb();
            FillProdOrderRows();
            _displayProdOrdersDt.AcceptChanges();
        }

        private void AddOrderRow(OrderRow order)
        {
            var newRow = _displayOrdersDt.NewRow();
            var fi = typeof(OrderRow).GetFields(BindingFlags.Public | BindingFlags.Instance);
            foreach (var info in fi) newRow[info.Name] = info.GetValue(order);

            _displayOrdersDt.Rows.Add(newRow);
        }

        private void AddProdOrderRow(ProdOrderRow order)
        {
            var newRow = _displayProdOrdersDt.NewRow();
            var fi = typeof(ProdOrderRow).GetFields(BindingFlags.Public | BindingFlags.Instance);
            foreach (var info in fi) newRow[info.Name] = info.GetValue(order);

            _displayProdOrdersDt.Rows.Add(newRow);
        }

        private static decimal CalculateKashirM2(int length, int width, int produceCount)
        {
            var result = (decimal)length * width * produceCount / 1000000;
            result = decimal.Round(result, 3, MidpointRounding.AwayFromZero);

            return result;
        }

        private static decimal CalculateKashirSheetsCount(int produceCount)
        {
            var result = (decimal)produceCount / 3;
            result = decimal.Round(result, 3, MidpointRounding.AwayFromZero);

            return result;
        }

        private static decimal CalculateM3(decimal thickness, int length, int width, int produceCount)
        {
            var result = thickness * length * (width + M3WidthConst) * produceCount / 1000000000;
            result = decimal.Round(result, 3, MidpointRounding.AwayFromZero);

            return result;
        }

        private static decimal CalculatePalletCount(decimal thickness, int width, int produceCount)
        {
            var result = thickness * width * produceCount / 1000000;
            result = decimal.Round(result, 3, MidpointRounding.AwayFromZero);

            return result;
        }

        private static decimal CalculateProfileM2(int length, int produceCount)
        {
            var result = (decimal)(length * produceCount) / 1000;
            result = decimal.Round(result, 3, MidpointRounding.AwayFromZero);

            return result;
        }

        private static decimal CalculateProfileSheetsCount(int length, int width, int produceCount)
        {
            decimal result = 0;
            switch (length)
            {
                case 2850:
                    result = produceCount / Math.Round(2070 / (width + M3WidthConst), 3);
                    break;

                case 2800:
                    switch (width)
                    {
                        case 2070:
                            result = produceCount / Math.Round(2070 / (width + M3WidthConst), 3);
                            break;

                        case 550:
                            result = produceCount / Math.Round(550 / (width + M3WidthConst), 3);
                            break;
                    }
                    break;

                case 2620:
                    result = produceCount / Math.Round(2070 / (width + M3WidthConst), 3);
                    break;

                case 2500:
                    result = produceCount / Math.Round(2070 / (width + M3WidthConst), 3);
                    break;

                case 2070:
                    result = produceCount / Math.Round(2800 / (width + M3WidthConst), 3);
                    break;
            }

            result = decimal.Round(result, 3, MidpointRounding.AwayFromZero);

            return result;
        }

        private void FillOrderRows(ClientsManager clientsManager, bool showZero)
        {
            for (var i = 0; i < _dbOrdersDt.Rows.Count; i++)
            {
                var decorOrderId = Convert.ToInt32(_dbOrdersDt.Rows[i]["DecorOrderID"]);
                var decorId = Convert.ToInt32(_dbOrdersDt.Rows[i]["DecorID"]);
                var colorId = Convert.ToInt32(_dbOrdersDt.Rows[i]["ColorID"]);
                var factCount = GetFactCount(decorOrderId, decorId);
                var decorName = GetDecorName(decorId);
                var colorName = GetColorName(colorId);
                decimal thickness = 0;
                var length = 0;
                var width = 0;
                var planCount = 0;
                decimal sheetsCount;
                decimal m2;
                var clientName = clientsManager.GetClientName(Convert.ToInt32(_dbOrdersDt.Rows[i]["ClientID"]));

                if (_dbOrdersDt.Rows[i]["Thickness"] != DBNull.Value)
                    thickness = Convert.ToDecimal(_dbOrdersDt.Rows[i]["Thickness"]);
                if (_dbOrdersDt.Rows[i]["Length"] != DBNull.Value)
                    length = Convert.ToInt32(_dbOrdersDt.Rows[i]["Length"]);
                if (_dbOrdersDt.Rows[i]["Width"] != DBNull.Value)
                    width = Convert.ToInt32(_dbOrdersDt.Rows[i]["Width"]);
                if (_dbOrdersDt.Rows[i]["Count"] != DBNull.Value)
                    planCount = Convert.ToInt32(_dbOrdersDt.Rows[i]["Count"]);

                var produceCount = planCount - factCount;

                if (!showZero &&
                    produceCount <= 0)
                    continue;

                if (planCount <= 2)
                    continue;

                var techStoreSubGroupId = 0;

                if (_dbOrdersDt.Rows[i]["TechStoreSubGroupID"] != DBNull.Value)
                    techStoreSubGroupId = Convert.ToInt32(_dbOrdersDt.Rows[i]["TechStoreSubGroupID"]);

                if (techStoreSubGroupId == Convert.ToInt32(TechStoreSubGroups.Kashir))
                    sheetsCount = CalculateKashirSheetsCount(produceCount);
                else
                    sheetsCount = CalculateProfileSheetsCount(length, width, produceCount);

                if (techStoreSubGroupId == Convert.ToInt32(TechStoreSubGroups.Kashir))
                    m2 = CalculateKashirM2(length, width, produceCount);
                else
                    m2 = CalculateProfileM2(length, produceCount);

                var m3 = CalculateM3(thickness, length, width, produceCount);

                bool toProduce = IsOrderToProduce(decorOrderId);

                var newRow = new OrderRow
                {
                    DecorName = decorName,
                    ColorName = colorName,
                    Thickness = thickness,
                    DecorOrderId = decorOrderId,
                    Length = length,
                    Width = width,
                    PlanCount = planCount,
                    FactCount = factCount,
                    ProduceCount = produceCount,
                    M3 = m3,
                    SheetsCount = sheetsCount,
                    M2 = m2,
                    ClientName = clientName,
                    ToProduce = toProduce
                };
                AddOrderRow(newRow);
            }
        }

        private void FillProdOrderRows()
        {
            for (var i = 0; i < _dbProdOrdersDt.Rows.Count; i++)
            {
                var decorOrderId = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["DecorOrderID"]);
                var decorId = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["DecorID"]);
                var colorId = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["ColorID"]);
                var decorName = GetDecorName(decorId);
                var colorName = GetColorName(colorId);
                decimal thickness = 0;
                var length = 0;
                var width = 0;
                var produceCount = 0;
                var id = 0;
                var priority = 0;
                var facingOption = 0;
                var millingOption = 0;
                var palletNumber = 0;
                decimal palletCount;
                decimal m2;

                if (_dbProdOrdersDt.Rows[i]["Thickness"] != DBNull.Value)
                    thickness = Convert.ToDecimal(_dbProdOrdersDt.Rows[i]["Thickness"]);
                if (_dbProdOrdersDt.Rows[i]["Length"] != DBNull.Value)
                    length = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["Length"]);
                if (_dbProdOrdersDt.Rows[i]["Width"] != DBNull.Value)
                    width = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["Width"]);
                if (_dbProdOrdersDt.Rows[i]["ProduceCount"] != DBNull.Value)
                    produceCount = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["ProduceCount"]);
                if (_dbProdOrdersDt.Rows[i]["id"] != DBNull.Value)
                    id = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["id"]);
                if (_dbProdOrdersDt.Rows[i]["priority"] != DBNull.Value)
                    priority = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["priority"]);
                if (_dbProdOrdersDt.Rows[i]["millingOption"] != DBNull.Value)
                    millingOption = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["millingOption"]);
                if (_dbProdOrdersDt.Rows[i]["facingOption"] != DBNull.Value)
                    facingOption = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["facingOption"]);

                var techStoreSubGroupId = 0;

                if (_dbProdOrdersDt.Rows[i]["TechStoreSubGroupID"] != DBNull.Value)
                    techStoreSubGroupId = Convert.ToInt32(_dbProdOrdersDt.Rows[i]["TechStoreSubGroupID"]);

                if (techStoreSubGroupId == Convert.ToInt32(TechStoreSubGroups.Kashir))
                    m2 = CalculateKashirM2(length, width, produceCount);
                else
                    m2 = CalculateProfileM2(length, produceCount);

                palletCount = CalculatePalletCount(thickness, width, produceCount);

                var newRow = new ProdOrderRow
                {
                    DecorName = decorName,
                    ColorName = colorName,
                    Thickness = thickness,
                    DecorOrderId = decorOrderId,
                    Length = length,
                    Width = width,
                    ProduceCount = produceCount,
                    Id = id,
                    DecorId = decorId,
                    Priority = priority,
                    M2 = m2,
                    PalletCount = palletCount,
                    FacingOption = facingOption,
                    MillingOption = millingOption,
                    PalletNumber = palletNumber
                };
                AddProdOrderRow(newRow);
            }
        }

        public readonly string[] ColNamesHide = { "Width", "Length", "DecorOrderId", "Thickness", "Id", "FacingOption", "MillingOption" };

        private string GetColorName(int id)
        {
            var name = "no_name";

            var rows = _colorsDt.Select("ColorID=" + id);
            if (rows.Any())
                name = rows[0]["ColorName"].ToString();
            return name;
        }

        private bool IsOrderToProduce(int id)
        {
            var rows = _moldingsSectionToProduceDt.Select("DecorOrderID=" + id);
            return rows.Any();
        }

        private void GetColorsDt()
        {
            _colorsDt = new DataTable();
            _colorsDt.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            _colorsDt.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            _colorsDt.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            const string selectCommand = @"SELECT TechStoreID, TechStoreName, Cvet FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE TechStoreGroupID = 11)
                ORDER BY TechStoreName";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString))
            {
                using (var dt = new DataTable())
                {
                    da.Fill(dt);
                    {
                        var NewRow = _colorsDt.NewRow();
                        NewRow["ColorID"] = -1;
                        NewRow["ColorName"] = "-";
                        NewRow["Cvet"] = "000";
                        _colorsDt.Rows.Add(NewRow);
                    }
                    {
                        var NewRow = _colorsDt.NewRow();
                        NewRow["ColorID"] = 0;
                        NewRow["ColorName"] = "на выбор";
                        NewRow["Cvet"] = "0000000";
                        _colorsDt.Rows.Add(NewRow);
                    }
                    for (var i = 0; i < dt.Rows.Count; i++)
                    {
                        var NewRow = _colorsDt.NewRow();
                        NewRow["ColorID"] = Convert.ToInt64(dt.Rows[i]["TechStoreID"]);
                        NewRow["ColorName"] = dt.Rows[i]["TechStoreName"].ToString();
                        NewRow["Cvet"] = dt.Rows[i]["Cvet"].ToString();
                        _colorsDt.Rows.Add(NewRow);
                    }
                }
            }
        }

        private void GetDecorDt()
        {
            _decorDt = new DataTable();
            const string selectCommand =
                @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID ORDER BY TechStoreName";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString))
            {
                da.Fill(_decorDt);
            }
        }

        private void GetMillingMachinesDt()
        {
            _millingMachinesDt = new DataTable();
            const string selectCommand =
                @"SELECT MachineID, MachineName FROM Machines WHERE SubSectorID=2 ORDER BY MachineName";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString))
            {
                da.Fill(_millingMachinesDt);
            }
            DataRow newRow = _millingMachinesDt.NewRow();
            newRow["MachineID"] = 0;
            newRow["MachineName"] = "-";
            _millingMachinesDt.Rows.InsertAt(newRow, 0);
        }

        private void GetFacingMachinesDt()
        {
            _facingMachinesDt = new DataTable();
            const string selectCommand =
                @"SELECT MachineID, MachineName FROM Machines WHERE SubSectorID=4 ORDER BY MachineName";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString))
            {
                da.Fill(_facingMachinesDt);
            }
            DataRow newRow = _facingMachinesDt.NewRow();
            newRow["MachineID"] = 0;
            newRow["MachineName"] = "-";
            _facingMachinesDt.Rows.InsertAt(newRow, 0);
        }

        public void FilterMillingMachines(int decorId)
        {
            DataTable dt = new DataTable();

            string selectCommand =
                $@"SELECT MachineID, MachinesOperationName FROM MachinesOperations where MachinesOperationID in (select MachinesOperationID from TechCatalogOperationsDetail
                where TechCatalogOperationsGroupID in (select TechCatalogOperationsGroupID from TechCatalogOperationsGroups where TechStoreID = {decorId}))";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString))
            {
                da.Fill(dt);
            }
            _millingMachinesBs.RemoveFilter();
            string filter = "0,";
            foreach (DataRow item in dt.Rows)
                filter += item["MachineID"] + ",";
            if (filter.Length > 0)
                filter = "MachineID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            _millingMachinesBs.Filter = filter;
            dt.Dispose();
        }

        public void FilterFacingMachines(int decorId)
        {
            DataTable dt = new DataTable();

            string selectCommand =
                $@"SELECT MachineID, MachinesOperationName FROM MachinesOperations where MachinesOperationID in (select MachinesOperationID from TechCatalogOperationsDetail
                where TechCatalogOperationsGroupID in (select TechCatalogOperationsGroupID from TechCatalogOperationsGroups where TechStoreID = {decorId}))";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString))
            {
                da.Fill(dt);
            }
            _facingMachinesBs.RemoveFilter();
            string filter = "0,";
            foreach (DataRow item in dt.Rows)
                filter += item["MachineID"] + ",";
            if (filter.Length > 0)
                filter = "MachineID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            _facingMachinesBs.Filter = filter;
            dt.Dispose();
        }

        private string GetDecorName(int id)
        {
            var name = "no_name";

            var rows = _decorDt.Select("DecorID=" + id);
            if (rows.Any())
                name = rows[0]["Name"].ToString();
            return name;
        }

        private int GetFactCount(int id, int decorId)
        {
            var factCount = 0;

            var rows = _decorAssignments.Select("DecorOrderID=" + id + " AND TechStoreID2=" + decorId);
            if (rows.Any() && rows[0]["FactCount"] != DBNull.Value)
                factCount = Convert.ToInt32(rows[0]["FactCount"]);
            return factCount;
        }

        private int GetMaxPriority()
        {
            int maxPriority = 0;

            if (_dbProdOrdersDt.Rows.Count > 0)
                maxPriority = Convert.ToInt32(_dbProdOrdersDt.Compute("MAX(Priority)", ""));

            return maxPriority;
        }

        private int GetMinPriority()
        {
            int minPriority = 0;

            if (_dbProdOrdersDt.Rows.Count > 0)
                minPriority = Convert.ToInt32(_dbProdOrdersDt.Compute("MIN(Priority)", ""));

            return minPriority;
        }

        public bool DownPriority(int id, int priority)
        {
            if (priority == GetMinPriority())
                return false;

            int prevPriority = priority - 1;

            var rows = _displayProdOrdersDt.Select("priority=" + prevPriority);
            if (rows.Any() && rows[0]["priority"] != DBNull.Value)
                rows[0]["priority"] = priority;

            rows = _displayProdOrdersDt.Select("id=" + id);
            if (rows.Any() && rows[0]["priority"] != DBNull.Value)
                rows[0]["priority"] = prevPriority;

            return true;
        }

        public bool UpPriority(int id, int priority)
        {
            if (priority == GetMaxPriority())
                return false;

            int nextPriority = priority + 1;

            var rows = _displayProdOrdersDt.Select("priority=" + nextPriority);
            if (rows.Any() && rows[0]["priority"] != DBNull.Value)
                rows[0]["priority"] = priority;

            rows = _displayProdOrdersDt.Select("id=" + id);
            if (rows.Any() && rows[0]["priority"] != DBNull.Value)
                rows[0]["priority"] = nextPriority;

            return true;
        }

        private void GetOrdersFromDb(bool onProd, bool inProd, bool onApproval, bool allClients, int clientId = -1)
        {
            var produceFilter = "";
            var agreementFilter = "";
            var clientFilter = "";
            var subGroupFilter = $@" AND (T.TechStoreSubGroupID={Convert.ToInt32(TechStoreSubGroups.Kashir)}
                OR T.TechStoreSubGroupID={Convert.ToInt32(TechStoreSubGroups.MilledAssembledProfile)}
                OR T.TechStoreSubGroupID={Convert.ToInt32(TechStoreSubGroups.MilledProfile)}
                OR T.TechStoreSubGroupID={Convert.ToInt32(TechStoreSubGroups.ShroudedAssembledProfile)}
                OR T.TechStoreSubGroupID={Convert.ToInt32(TechStoreSubGroups.ShroudedProfile)})";
            if (onProd)
                produceFilter = "WHERE ProfilProductionStatusID = 3";
            if (inProd)
                produceFilter = "WHERE ProfilProductionStatusID = 2";
            if (onApproval)
                agreementFilter = " AND AgreementStatusID=3";
            if (!allClients)
                clientFilter = " AND clientId = " + clientId;

            var selectCommand =
                $@"SELECT T.TechStoreSubGroupID,T.Thickness,Mega.ClientID,Main.MegaOrderID,D.DecorOrderID,D.MainOrderID,D.ProductID,D.DecorID,D.ColorID,D.PatinaID,
                D.InsetTypeID,D.InsetColorID,D.Length,D.Height,D.Width,D.Count,D.DecorConfigID,D.ItemWeight,D.Weight FROM NewDecorOrders as D
                INNER JOIN NewMainOrders as Main ON D.MainOrderID=Main.MainOrderID AND Main.MainOrderID IN (SELECT MainOrderID FROM NewMainOrders {produceFilter})
                INNER JOIN NewMegaOrders as Mega ON Main.MegaOrderID=Mega.MegaOrderID {clientFilter} {agreementFilter}
                INNER JOIN infiniu2_catalog.dbo.TechStore as T ON D.DecorID=T.TechStoreID {subGroupFilter} ORDER BY T.Thickness";

            _dbOrdersDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand,
                       ConnectionStrings.MarketingOrdersConnectionString))
            {
                da.Fill(_dbOrdersDt);
            }

            selectCommand =
                $@"SELECT DecorAssignmentID,ClientID,MegaOrderID,MainOrderID,DecorOrderID,TechStoreID2,PlanCount,FactCount FROM DecorAssignments
                WHERE MainOrderID IN (SELECT MainOrderID FROM infiniu2_marketingorders.dbo.NewMainOrders {produceFilter})";

            _decorAssignments.Clear();
            using (var da = new SqlDataAdapter(selectCommand,
                       ConnectionStrings.StorageConnectionString))
            {
                da.Fill(_decorAssignments);
            }
        }

        private void GetProdOrdersFromDb()
        {
            var selectCommand =
                @"SELECT M.*,T.TechStoreSubGroupID,T.Thickness,Mega.ClientID,Main.MegaOrderID,D.MainOrderID,D.ProductID,D.DecorID,D.ColorID,D.PatinaID,
                D.InsetTypeID,D.InsetColorID,D.Length,D.Height,D.Width,D.Count,D.DecorConfigID,D.ItemWeight,D.Weight FROM infiniu2_storage.dbo.MoldingsSectionToProduce as M
                INNER JOIN NewDecorOrders as D ON M.DecorOrderID=D.DecorOrderID
                INNER JOIN NewMainOrders as Main ON D.MainOrderID=Main.MainOrderID
                INNER JOIN NewMegaOrders as Mega ON Main.MegaOrderID=Mega.MegaOrderID
                INNER JOIN infiniu2_catalog.dbo.TechStore as T ON D.DecorID=T.TechStoreID ORDER BY M.Priority";

            _dbProdOrdersDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand,
                       ConnectionStrings.MarketingOrdersConnectionString))
            {
                da.Fill(_dbProdOrdersDt);
            }
        }

        private void GetSavedOrdersFromDb()
        {
            var selectCommand =
                @"SELECT * FROM MoldingsSectionToProduce";

            _moldingsSectionToProduceDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand,
                       ConnectionStrings.StorageConnectionString))
            {
                da.Fill(_moldingsSectionToProduceDt);
            }
        }

        private void GetUsersDt()
        {
            _usersDt = new DataTable();
            const string selectCommand = @"SELECT UserID, ShortName FROM Users";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.UsersConnectionString))
            {
                da.Fill(_usersDt);
            }
        }

        private void GroupByThickness()
        {
            OrdersDs.Tables.Clear();
            DataTable itemsDt;

            using (var dv = new DataView(_displayOrdersDt))
            {
                dv.Sort = "Thickness, DecorName, planCount";
                itemsDt = dv.ToTable(true, "Thickness");
            }

            for (var i = 0; i < itemsDt.Rows.Count; i++)
            {
                var thickness = Convert.ToInt32(itemsDt.Rows[i]["Thickness"]);
                var rows = _displayOrdersDt.Select("Thickness=" + thickness);
                var dt = rows.CopyToDataTable();
                OrdersDs.Tables.Add(dt);
            }
        }

        private void UnGroup()
        {
            OrdersDs.Tables.Clear();

            DataTable itemsDt;

            using (var dv = new DataView(_displayOrdersDt))
            {
                dv.Sort = "Thickness, DecorName, ColorName, planCount";
                itemsDt = dv.ToTable(false);
            }
            OrdersDs.Tables.Add(itemsDt);
        }

        public DataGridViewComboBoxColumn MillingOptionColumn
        {
            get
            {
                DataGridViewComboBoxColumn column = new DataGridViewComboBoxColumn
                {
                    Name = "MillingOptionColumn",
                    HeaderText = "Вариант фрезеровки",
                    DataPropertyName = "MillingOption",
                    DataSource = _millingMachinesBs,
                    ValueMember = "MachineID",
                    DisplayMember = "MachineName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return column;
            }
        }

        public DataGridViewComboBoxColumn FacingOptionColumn
        {
            get
            {
                DataGridViewComboBoxColumn column = new DataGridViewComboBoxColumn
                {
                    Name = "FacingOptionColumn",
                    HeaderText = "Вариант облицовки",
                    DataPropertyName = "FacingOption",
                    DataSource = _facingMachinesBs,
                    ValueMember = "MachineID",
                    DisplayMember = "MachineName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return column;
            }
        }

        private struct OrderRow
        {
            public string DecorName;
            public string ColorName;
            public decimal Thickness;
            public int Length;
            public int Width;
            public int PlanCount;
            public int ProduceCount;
            public decimal M3;
            public decimal SheetsCount;
            public decimal M2;
            public string ClientName;
            public bool ToProduce;
            public int DecorOrderId;
            public int FactCount;
        }

        private struct ProdOrderRow
        {
            public string DecorName;
            public string ColorName;
            public int ProduceCount;
            public decimal M2;
            public int Priority;
            public long MillingOption;
            public decimal PalletCount;
            public int PalletNumber;
            public long FacingOption;
            public int DecorOrderId;
            public int DecorId;
            public int Id;
            public decimal Thickness;
            public int Length;
            public int Width;
        }
    }

    public class PermissionsManager
    {
        private readonly DataTable rolePermissionsDt;

        public PermissionsManager()
        {
            rolePermissionsDt = new DataTable();
        }

        public void GetPermissions(int UserID, string FormName)
        {
            using (var da = new SqlDataAdapter($@"SELECT * FROM UserRoles WHERE UserID = {UserID}
                AND RoleID IN (SELECT RoleID FROM Roles WHERE ModuleID IN (SELECT ModuleID FROM Modules WHERE FormName = '{FormName}'))",
                       ConnectionStrings.UsersConnectionString))
            {
                da.Fill(rolePermissionsDt);
            }
        }

        public bool PermissionGranted(int RoleID)
        {
            var rows = rolePermissionsDt.Select("RoleID = " + RoleID);
            return rows.Any();
        }
    }
}