using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.LoadCalculations
{
    public enum OrderStatus
    {
        NotConfirmed,
        ForAgreed,
        Agreed,
        OnProduction,
        InProduction
    }

    public class LoadCalculations
    {
        private readonly DataTable _clientsDt;

        private readonly DataTable _clientsOrdersDt;

        private readonly DataTable _clientsNotConfirmedDt;
        private readonly DataTable _clientsForAgreedDt;
        private readonly DataTable _clientsAgreedDt;
        private readonly DataTable _clientsOnProductionDt;
        private readonly DataTable _clientsInProductionDt;

        private readonly DataTable _frontsOrdersDt;
        private readonly DataTable _decorOrdersDt;
        private readonly DataTable _groupbyClientDt;
        private readonly DataTable _groupbyMachinesDt;
        private readonly DataTable _groupbySectorsDt;
        private readonly DataTable _loadCalculationsDt;
        private readonly DataTable _machinePriorityDt;
        private readonly DataTable _machineRatesDt;
        private readonly DataTable _machinesDt;
        private readonly DataTable _techstoreDt;
        private readonly DataTable _distDecorDt;
        private readonly DataTable _staffingDt;


        public LoadCalculations()
        {
            _clientsDt = new DataTable();
            _frontsOrdersDt = new DataTable();
            _decorOrdersDt = new DataTable();
            _machinePriorityDt = new DataTable();
            _machineRatesDt = new DataTable();
            _machinesDt = new DataTable();
            _techstoreDt = new DataTable();
            _distDecorDt = new DataTable();
            _staffingDt = new DataTable();

            //_groupbyClientDt = new DataTable();
            //_groupbyClientDt.Columns.Add(new DataColumn("id", typeof(int)) { AutoIncrement = true });
            //_groupbyClientDt.Columns.Add("machineId", typeof(int));
            //_groupbyClientDt.Columns.Add("sectorId", typeof(int));
            //_groupbyClientDt.Columns.Add("clientId", typeof(int));
            //_groupbyClientDt.Columns.Add("orderNumber", typeof(int));
            //_groupbyClientDt.Columns.Add("total1", typeof(decimal));
            //_groupbyClientDt.Columns.Add("total2", typeof(decimal));
            //_groupbyClientDt.Columns.Add("total3", typeof(decimal));
            //_groupbyClientDt.Columns.Add("total4", typeof(decimal));

            _groupbyMachinesDt = new DataTable();
            _groupbyMachinesDt.Columns.Add(new DataColumn("id", typeof(int)) { AutoIncrement = true });
            _groupbyMachinesDt.Columns.Add("machineId", typeof(int));
            _groupbyMachinesDt.Columns.Add("sectorId", typeof(int));
            _groupbyMachinesDt.Columns.Add("clientId", typeof(int));
            _groupbyMachinesDt.Columns.Add("orderNumber", typeof(int));
            _groupbyMachinesDt.Columns.Add("orderStatus", typeof(int));
            _groupbyMachinesDt.Columns.Add("clientName", typeof(string));
            _groupbyMachinesDt.Columns.Add("total1", typeof(decimal));
            _groupbyMachinesDt.Columns.Add("total2", typeof(decimal));
            _groupbyMachinesDt.Columns.Add("total3", typeof(decimal));
            _groupbyMachinesDt.Columns.Add("total4", typeof(decimal));

            _distDecorDt = new DataTable();
            _distDecorDt.Columns.Add("decorName", typeof(string));
            _distDecorDt.Columns.Add("sectorId", typeof(int));
            _distDecorDt.Columns.Add("machineId", typeof(int));
            _distDecorDt.Columns.Add("mainOrderId", typeof(int));

            _groupbyClientDt = _groupbyMachinesDt.Clone();

            _groupbySectorsDt = new DataTable();
            _groupbySectorsDt.Columns.Add(new DataColumn("id", typeof(int)) { AutoIncrement = true });
            _groupbySectorsDt.Columns.Add("sectorId", typeof(int));
            _groupbySectorsDt.Columns.Add("total1", typeof(decimal));
            _groupbySectorsDt.Columns.Add("total2", typeof(decimal));
            _groupbySectorsDt.Columns.Add("total3", typeof(decimal));
            _groupbySectorsDt.Columns.Add("total4", typeof(decimal));
            _groupbySectorsDt.Columns.Add("sumtotal", typeof(decimal));

            _loadCalculationsDt = new DataTable();
            _loadCalculationsDt.Columns.Add(new DataColumn("id", typeof(int)) { AutoIncrement = true });
            _loadCalculationsDt.Columns.Add("userId", typeof(int));
            _loadCalculationsDt.Columns.Add("machineId", typeof(int));
            _loadCalculationsDt.Columns.Add("sectorId", typeof(int));
            _loadCalculationsDt.Columns.Add("clientId", typeof(int));
            _loadCalculationsDt.Columns.Add("orderNumber", typeof(int));
            _loadCalculationsDt.Columns.Add("orderStatus", typeof(int));
            _loadCalculationsDt.Columns.Add("count", typeof(int));
            _loadCalculationsDt.Columns.Add("machineName", typeof(string));
            _loadCalculationsDt.Columns.Add("decorName", typeof(string));
            _loadCalculationsDt.Columns.Add("rate1", typeof(decimal));
            _loadCalculationsDt.Columns.Add("rate2", typeof(decimal));
            _loadCalculationsDt.Columns.Add("rate3", typeof(decimal));
            _loadCalculationsDt.Columns.Add("rate4", typeof(decimal));
            _loadCalculationsDt.Columns.Add("total1", typeof(decimal));
            _loadCalculationsDt.Columns.Add("total2", typeof(decimal));
            _loadCalculationsDt.Columns.Add("total3", typeof(decimal));
            _loadCalculationsDt.Columns.Add("total4", typeof(decimal));
            _loadCalculationsDt.Columns.Add("sumtotal", typeof(decimal));

            _clientsAgreedDt = new DataTable();
            _clientsAgreedDt.Columns.Add("clientName", typeof(string));
            _clientsAgreedDt.Columns.Add("orderNumber", typeof(int));
            _clientsAgreedDt.Columns.Add("OrderDate", typeof(string));
            _clientsAgreedDt.Columns.Add("sumTotal", typeof(decimal));

            _clientsOrdersDt = new DataTable();
            _clientsOrdersDt.Columns.Add("clientId", typeof(int));
            _clientsOrdersDt.Columns.Add("clientName", typeof(string));
            _clientsOrdersDt.Columns.Add("orderNumber", typeof(int));
            _clientsOrdersDt.Columns.Add("orderStatus", typeof(int));
            _clientsOrdersDt.Columns.Add("OrderDate", typeof(string));

            _clientsNotConfirmedDt = _clientsAgreedDt.Clone();
            _clientsForAgreedDt = _clientsAgreedDt.Clone();
            _clientsOnProductionDt = _clientsAgreedDt.Clone();
            _clientsInProductionDt = _clientsAgreedDt.Clone();

            ClientsNotConfirmedList = new BindingSource { DataSource = _clientsNotConfirmedDt };
            ClientsForAgreedList = new BindingSource { DataSource = _clientsForAgreedDt };
            ClientsAgreedList = new BindingSource { DataSource = _clientsAgreedDt };
            ClientsOnProductionList = new BindingSource { DataSource = _clientsOnProductionDt };
            ClientsInProductionList = new BindingSource { DataSource = _clientsInProductionDt };

            GetClients();
            GetMachines();
            GetTechStore();
            GetMachinePriority();
            GetMachineRates();
            GetStaffing();
        }

        public BindingSource ClientsNotConfirmedList
        {
            get;
        }

        public BindingSource ClientsForAgreedList
        {
            get;
        }

        public BindingSource ClientsAgreedList
        {
            get;
        }

        public BindingSource ClientsOnProductionList
        {
            get;
        }

        public BindingSource ClientsInProductionList
        {
            get;
        }

        private decimal SumRatesByRank1(int sectorId, int rank)
        {
            var filter = $"sectorId={sectorId} and rank={rank}";
            if (rank == 4)
                filter = $"sectorId={sectorId} and rank>={rank}";

            var rows = _staffingDt.Select(filter);
            if (!rows.Any())
                return 0;

            decimal sumRates = 0;

            foreach (var row in rows)
                sumRates += Convert.ToDecimal(row["rate"]);

            return sumRates;
        }

        private decimal SumRatesByRank2(int sectorId, int rank)
        {

            //if (!_staffingDt
            //        .AsEnumerable().Any(row => row.Field<long>("sectorId") == sectorId &&
            //                                   row.Field<int>("rank") == rank))
            //    return 0;

            var sumRates = _staffingDt.AsEnumerable()
                .Where(row => row.Field<long>("sectorId") == sectorId &&
                              row.Field<int>("rank") == rank)
                .Sum(x => x.Field<decimal>("rate"));

            return sumRates;
        }

        private string ClientName(int id)
        {
            var name = "";

            var rows = _clientsDt.Select($"clientId={id}");
            if (rows.Any())
                name = rows[0]["ClientName"].ToString();
            return name;
        }

        private string MachineName(int id)
        {
            var name = "";

            var rows = _machinesDt.Select($"machineId={id}");
            if (rows.Any())
                name = rows[0]["machineName"].ToString();
            return name;
        }

        private string TechStorename(int id)
        {
            var name = "";

            var rows = _techstoreDt.Select($"techstoreId={id}");
            if (rows.Any())
                name = rows[0]["techstorename"].ToString();
            return name;
        }

        private string SectorName(int id)
        {
            var name = "";

            var rows = _machinesDt.Select($"sectorId={id}");
            if (rows.Any())
                name = rows[0]["sectorName"].ToString();
            return name;
        }

        private string UserName(int id)
        {
            var name = "";

            var rows = _staffingDt.Select($"userId={id}");
            if (rows.Any())
                name = rows[0]["ShortName"].ToString();
            return name;
        }

        public int RankCoef
        {
            get;
            set;
        } = 1;

        public decimal AgreedSumRank
        {
            get;
            private set;
        } = 0;

        public decimal NotConfirmedSumRank
        {
            get;
            private set;
        } = 0;

        public decimal ForAgreedSumRank
        {
            get;
            private set;
        } = 0;

        public decimal OnProductionSumRank
        {
            get;
            private set;
        } = 0;

        public decimal InProductionSumRank
        {
            get;
            private set;
        } = 0;

        public List<Machine> GroupByMachines(int sectorId)
        {
            _groupbyMachinesDt.Clear();

            var machinesList = new List<Machine>();

            var distMachines = _loadCalculationsDt.AsEnumerable()
                .Where(row => row.Field<int>("sectorId") == sectorId)
                .GroupBy(g => new { machineId = g.Field<int>("machineId") })
                .Select(s => s.Key.machineId).ToList();

            var distClients = _loadCalculationsDt.AsEnumerable()
                .GroupBy(g => new
                {
                    clientId = g.Field<int>("clientId"),
                    orderNumber = g.Field<int>("orderNumber"),
                    orderStatus = g.Field<int>("orderStatus")
                })
                .Select(s => new Client(ClientName(s.Key.clientId),
                    s.Key.clientId,
                    s.Key.orderNumber,
                    s.Key.orderStatus)).OrderBy(x => x.Name).ToList();

            FillClientsDt();

            foreach (var machineId in distMachines)
            {
                _groupbyClientDt.Clear();

                foreach (var item in distClients)
                    AddClientRow(0, sectorId, item.ClientId, item.OrderNumber, item.OrderStatus, 0, 0, 0, 0);

                var clients = _loadCalculationsDt.AsEnumerable()
                    .Where(row => row.Field<int>("sectorId") == sectorId && row.Field<int>("machineId") == machineId)
                    .GroupBy(g => new
                    {
                        clientId = g.Field<int>("clientId"),
                        orderNumber = g.Field<int>("orderNumber"),
                        orderStatus = g.Field<int>("orderStatus")
                    })
                    .Select(s => new Client(ClientName(s.Key.clientId),
                        s.Key.clientId,
                        s.Key.orderNumber,
                        s.Key.orderStatus)).OrderBy(x => x.Name);

                var machine = new Machine();

                decimal sumTotal = 0;
                decimal t1 = 0;
                decimal t2 = 0;
                decimal t3 = 0;
                decimal t4 = 0;
                decimal total1 = 0;
                decimal total2 = 0;
                decimal total3 = 0;
                decimal total4 = 0;

                foreach (var item1 in clients)
                {
                    var clientId = item1.ClientId;
                    var orderNumber = item1.OrderNumber;
                    var orderStatus = item1.OrderStatus;

                    var details = _loadCalculationsDt.AsEnumerable()
                        .Where(row => row.Field<int>("sectorId") == sectorId
                                      && row.Field<int>("machineId") == machineId
                                      && row.Field<int>("clientId") == clientId
                                      && row.Field<int>("orderNumber") == orderNumber
                                      && row.Field<int>("orderStatus") == orderStatus)
                        .Select(s => new
                        {
                            total1 = decimal.Round(s.Field<decimal>("total1"), 3, MidpointRounding.AwayFromZero),
                            total2 = decimal.Round(s.Field<decimal>("total2"), 3, MidpointRounding.AwayFromZero),
                            total3 = decimal.Round(s.Field<decimal>("total3"), 3, MidpointRounding.AwayFromZero),
                            total4 = decimal.Round(s.Field<decimal>("total4"), 3, MidpointRounding.AwayFromZero)
                        });

                    foreach (var item2 in details)
                    {

                        EditClientRow(machineId, sectorId, clientId, orderNumber, orderStatus,
                            t1 = item2.total1,
                            t2 = item2.total2,
                            t3 = item2.total3,
                            t4 = item2.total4);

                        sumTotal += t1 + t2 + t3 + t4;
                        total1 += t1;
                        total2 += t2;
                        total3 += t3;
                        total4 += t4;
                    }
                }

                total1 = decimal.Round(total1, 3, MidpointRounding.AwayFromZero);
                total2 = decimal.Round(total2, 3, MidpointRounding.AwayFromZero);
                total3 = decimal.Round(total3, 3, MidpointRounding.AwayFromZero);
                total4 = decimal.Round(total4, 3, MidpointRounding.AwayFromZero);
                sumTotal = decimal.Round(sumTotal, 3, MidpointRounding.AwayFromZero);

                AddGroupByMachineRow(machineId, sectorId, 0, 0, total1, total2, total3, total4);

                machine.SumTotal = sumTotal;
                machine.Name = MachineName(machineId);
                machine.Id = machineId;

                machine.DataSource1 = _groupbyMachinesDt.AsEnumerable()
                    .Where(row => row.Field<int>("machineId") == machineId)
                    .Select(x => x).CopyToDataTable();

                if (_groupbyClientDt
                    .AsEnumerable().Any(row => row.Field<int>("orderStatus") == (int)OrderStatus.NotConfirmed))
                {
                    machine.DataSourceNotConfirmed =
                        new BindingSource
                        {
                            DataSource =
                                _groupbyClientDt.AsEnumerable()
                                    .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.NotConfirmed)
                                    .Select(x => x).CopyToDataTable()
                        };
                }
                if (_groupbyClientDt
                    .AsEnumerable().Any(row => row.Field<int>("orderStatus") == (int)OrderStatus.ForAgreed))
                {
                    machine.DataSourceForAgreed =
                        new BindingSource
                        {
                            DataSource =
                                _groupbyClientDt.AsEnumerable()
                                    .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.ForAgreed)
                                    .Select(x => x).CopyToDataTable()
                        };
                }

                if (_groupbyClientDt
                    .AsEnumerable().Any(row => row.Field<int>("orderStatus") == (int)OrderStatus.Agreed))
                {
                    machine.DataSourceAgreed =
                        new BindingSource
                        {
                            DataSource =
                                _groupbyClientDt.AsEnumerable()
                                    .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.Agreed)
                                    .Select(x => x).CopyToDataTable()
                        };
                }

                if (_groupbyClientDt
                    .AsEnumerable().Any(row => row.Field<int>("orderStatus") == (int)OrderStatus.OnProduction))
                {
                    machine.DataSourceOnProduction =
                        new BindingSource
                        {
                            DataSource =
                                _groupbyClientDt.AsEnumerable()
                                    .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.OnProduction)
                                    .Select(x => x).CopyToDataTable()
                        };
                }

                if (_groupbyClientDt
                    .AsEnumerable().Any(row => row.Field<int>("orderStatus") == (int)OrderStatus.InProduction))
                {
                    machine.DataSourceInProduction =
                        new BindingSource
                        {
                            DataSource =
                                _groupbyClientDt.AsEnumerable()
                                    .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.InProduction)
                                    .Select(x => x).CopyToDataTable()
                        };
                }

                machinesList.Add(machine);
            }

            return machinesList;
        }

        private void FillClientsDt()
        {
            _clientsNotConfirmedDt.Clear();
            _clientsForAgreedDt.Clear();
            _clientsAgreedDt.Clear();
            _clientsOnProductionDt.Clear();
            _clientsInProductionDt.Clear();

            var distClients = _clientsOrdersDt.AsEnumerable()
                .GroupBy(g => new
                {
                    clientId = g.Field<int>("clientId"),
                    orderNumber = g.Field<int>("orderNumber"),
                    orderStatus = g.Field<int>("orderStatus"),
                    OrderDate = g.Field<string>("OrderDate"),
                })
                .Select(s => new
                {
                    OrderDate = s.Key.OrderDate,
                    client = new Client(ClientName(s.Key.clientId),
                        s.Key.clientId,
                        s.Key.orderNumber,
                        s.Key.orderStatus)

                }).OrderBy(x => x.client.Name).ToList();

            foreach (var item in distClients)
            {
                decimal sumTotal = 0;
                DataRow newRow;
                switch (item.client.OrderStatus)
                {
                    case (int)OrderStatus.NotConfirmed:

                        newRow = _clientsNotConfirmedDt.NewRow();
                        newRow["clientName"] = item.client.Name;
                        newRow["orderNumber"] = item.client.OrderNumber;
                        newRow["OrderDate"] = item.OrderDate;

                        sumTotal = _loadCalculationsDt.AsEnumerable()
                            .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.NotConfirmed
                            && row.Field<int>("clientId") == item.client.ClientId
                            && row.Field<int>("OrderNumber") == item.client.OrderNumber)
                            .Select(x => x)
                            .Sum(s => s.Field<decimal>("total1") +
                                      s.Field<decimal>("total2") +
                                      s.Field<decimal>("total3") +
                                      s.Field<decimal>("total4"));
                        newRow["sumTotal"] = sumTotal;

                        _clientsNotConfirmedDt.Rows.Add(newRow);
                        break;
                    case (int)OrderStatus.ForAgreed:
                        newRow = _clientsForAgreedDt.NewRow();
                        newRow["clientName"] = item.client.Name;
                        newRow["orderNumber"] = item.client.OrderNumber;
                        newRow["OrderDate"] = item.OrderDate;

                        sumTotal = _loadCalculationsDt.AsEnumerable()
                            .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.ForAgreed
                                          && row.Field<int>("clientId") == item.client.ClientId
                                          && row.Field<int>("OrderNumber") == item.client.OrderNumber)
                            .Select(x => x)
                            .Sum(s => s.Field<decimal>("total1") +
                                      s.Field<decimal>("total2") +
                                      s.Field<decimal>("total3") +
                                      s.Field<decimal>("total4"));
                        newRow["sumTotal"] = sumTotal;

                        _clientsForAgreedDt.Rows.Add(newRow);
                        break;
                    case (int)OrderStatus.Agreed:
                        newRow = _clientsAgreedDt.NewRow();
                        newRow["clientName"] = item.client.Name;
                        newRow["orderNumber"] = item.client.OrderNumber;
                        newRow["OrderDate"] = item.OrderDate;

                        sumTotal = _loadCalculationsDt.AsEnumerable()
                            .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.Agreed
                                          && row.Field<int>("clientId") == item.client.ClientId
                                          && row.Field<int>("OrderNumber") == item.client.OrderNumber)
                            .Select(x => x)
                            .Sum(s => s.Field<decimal>("total1") +
                                      s.Field<decimal>("total2") +
                                      s.Field<decimal>("total3") +
                                      s.Field<decimal>("total4"));
                        newRow["sumTotal"] = sumTotal;

                        _clientsAgreedDt.Rows.Add(newRow);
                        break;
                    case (int)OrderStatus.OnProduction:
                        newRow = _clientsOnProductionDt.NewRow();
                        newRow["clientName"] = item.client.Name;
                        newRow["orderNumber"] = item.client.OrderNumber;
                        newRow["OrderDate"] = item.OrderDate;

                        sumTotal = _loadCalculationsDt.AsEnumerable()
                            .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.OnProduction
                                          && row.Field<int>("clientId") == item.client.ClientId
                                          && row.Field<int>("OrderNumber") == item.client.OrderNumber)
                            .Select(x => x)
                            .Sum(s => s.Field<decimal>("total1") +
                                      s.Field<decimal>("total2") +
                                      s.Field<decimal>("total3") +
                                      s.Field<decimal>("total4"));
                        newRow["sumTotal"] = sumTotal;

                        _clientsOnProductionDt.Rows.Add(newRow);
                        break;
                    case (int)OrderStatus.InProduction:
                        newRow = _clientsInProductionDt.NewRow();
                        newRow["clientName"] = item.client.Name;
                        newRow["orderNumber"] = item.client.OrderNumber;
                        newRow["OrderDate"] = item.OrderDate;

                        sumTotal = _loadCalculationsDt.AsEnumerable()
                            .Where(row => row.Field<int>("orderStatus") == (int)OrderStatus.InProduction
                                          && row.Field<int>("clientId") == item.client.ClientId
                                          && row.Field<int>("OrderNumber") == item.client.OrderNumber)
                            .Select(x => x)
                            .Sum(s => s.Field<decimal>("total1") +
                                      s.Field<decimal>("total2") +
                                      s.Field<decimal>("total3") +
                                      s.Field<decimal>("total4"));
                        newRow["sumTotal"] = sumTotal;

                        _clientsInProductionDt.Rows.Add(newRow);
                        break;
                    default:
                        break;
                }
            }
        }

        private void AddClientOrder(int clientId, string clientName, int orderNumber, int orderStatus, string orderDate)
        {
            var newRow = _clientsOrdersDt.NewRow();
            newRow["clientId"] = clientId;
            newRow["clientName"] = clientName;
            newRow["orderNumber"] = orderNumber;
            newRow["orderStatus"] = orderStatus;
            newRow["OrderDate"] = orderDate;
            _clientsOrdersDt.Rows.Add(newRow);
        }

        public List<Sector> SectorsList { get; set; } = new List<Sector>();

        public void CreateSectors(List<int> sectorsId)
        {
            SectorsList.Clear();

            foreach (var sectorId in sectorsId)
            {
                decimal rank1;
                decimal rank2;
                decimal rank3;
                decimal rank4;
                var sector = new Sector
                {
                    Name = SectorName(sectorId),
                    Id = sectorId,
                    Rank1 = rank1 = SumRatesByRank1(sectorId, 0) * RankCoef,
                    Rank2 = rank2 = SumRatesByRank1(sectorId, 2) * RankCoef,
                    Rank3 = rank3 = SumRatesByRank1(sectorId, 3) * RankCoef,
                    Rank4 = rank4 = SumRatesByRank1(sectorId, 4) * RankCoef,

                    //Total1 = total1 = item.total1,
                    //Total2 = total2 = item.total2,
                    //Total3 = total3 = item.total3,
                    //Total4 = total4 = item.total4,

                    SumRank = rank1 + rank2 + rank3 + rank4
                };
                SectorsList.Add(sector);
            }

            SectorsList = SectorsList.OrderByDescending(r => r.Name).ToList();
        }

        public void DeleteEmptySectors()
        {
            for (var i = SectorsList.Count - 1; i >= 0; i--)
            {
                if (!_loadCalculationsDt.Select("sectorId=" + SectorsList[i].Id).Any())
                {
                    SectorsList.RemoveAt(i);
                    continue;
                }

                //var sumTotal = SectorsList[i].Total1 + SectorsList[i].Total2 + SectorsList[i].Total3 + SectorsList[i].Total4;
                //SectorsList[i].SumTotal = sumTotal;
                //if (SectorsList[i].SumRank != 0)
                //    SectorsList[i].Percent = decimal.Round(sumTotal / SectorsList[i].SumRank * 100, 3, MidpointRounding.AwayFromZero);
            }
        }

        public void GroupBySectors(List<int> sectorsId)
        {
            _groupbySectorsDt.Clear();

            var groupTotal = _loadCalculationsDt.AsEnumerable()
                .Where(w => sectorsId.Contains(w.Field<int>("sectorId")))
                .GroupBy(r => new
                {
                    sectorId = r.Field<int>("sectorId")
                })
                .Select(g => new
                {
                    sectorId = g.Key.sectorId,
                    sectorName = SectorName(g.Key.sectorId),
                    total1 = g.Sum(r => decimal.Round(r.Field<decimal>("total1"), 3, MidpointRounding.AwayFromZero)),
                    total2 = g.Sum(r => decimal.Round(r.Field<decimal>("total2"), 3, MidpointRounding.AwayFromZero)),
                    total3 = g.Sum(r => decimal.Round(r.Field<decimal>("total3"), 3, MidpointRounding.AwayFromZero)),
                    total4 = g.Sum(r => decimal.Round(r.Field<decimal>("total4"), 3, MidpointRounding.AwayFromZero))
                }).OrderByDescending(x => x.sectorName);

            foreach (var item in groupTotal)
            {
                decimal total1 = 0;
                decimal total2 = 0;
                decimal total3 = 0;
                decimal total4 = 0;
                decimal rank1;
                decimal rank2;
                decimal rank3;
                decimal rank4;
                decimal sumRank;
                decimal sumTotal;

                var sectorIndex = SectorsList.FindIndex(s => s.Id == item.sectorId);
                SectorsList[sectorIndex].Name = item.sectorName;
                SectorsList[sectorIndex].Id = item.sectorId;
                SectorsList[sectorIndex].Rank1 = rank1 = SumRatesByRank1(item.sectorId, 0) * RankCoef;
                SectorsList[sectorIndex].Rank2 = rank2 = SumRatesByRank1(item.sectorId, 2) * RankCoef;
                SectorsList[sectorIndex].Rank3 = rank3 = SumRatesByRank1(item.sectorId, 3) * RankCoef;
                SectorsList[sectorIndex].Rank4 = rank4 = SumRatesByRank1(item.sectorId, 4) * RankCoef;

                SectorsList[sectorIndex].Total1 = total1 = item.total1;
                SectorsList[sectorIndex].Total2 = total2 = item.total2;
                SectorsList[sectorIndex].Total3 = total3 = item.total3;
                SectorsList[sectorIndex].Total4 = total4 = item.total4;

                SectorsList[sectorIndex].SumRank = sumRank = rank1 + rank2 + rank3 + rank4;
                SectorsList[sectorIndex].SumTotal = sumTotal = total1 + total2 + total3 + total4;
                if (sumRank != 0)
                    SectorsList[sectorIndex].Percent = decimal.Round(sumTotal / sumRank * 100, 3, MidpointRounding.AwayFromZero);
            }
        }

        public void ClearCalculations()
        {
            NotConfirmedSumRank = 0;
            ForAgreedSumRank = 0;
            AgreedSumRank = 0;
            OnProductionSumRank = 0;
            InProductionSumRank = 0;

            _clientsNotConfirmedDt.Clear();
            _clientsForAgreedDt.Clear();
            _clientsAgreedDt.Clear();
            _clientsOnProductionDt.Clear();
            _clientsInProductionDt.Clear();
            _clientsOrdersDt.Clear();
            _distDecorDt.Clear();
            _frontsOrdersDt.Clear();
            _decorOrdersDt.Clear();
            _loadCalculationsDt.Clear();
        }

        public void CalculateLoad(OrderStatus orderStatus)
        {
            if (_decorOrdersDt.Rows.Count == 0 && _frontsOrdersDt.Rows.Count == 0)
                return;

            for (var i = 0; i < _decorOrdersDt.Rows.Count; i++)
            {
                var mainOrderId = Convert.ToInt32(_decorOrdersDt.Rows[i]["mainOrderId"]);
                var clientId = Convert.ToInt32(_decorOrdersDt.Rows[i]["clientId"]);
                var OrderDate = string.Empty;
                var orderNumber = Convert.ToInt32(_decorOrdersDt.Rows[i]["orderNumber"]);
                var itemId = Convert.ToInt32(_decorOrdersDt.Rows[i]["productId"]);
                var decorId = Convert.ToInt32(_decorOrdersDt.Rows[i]["decorId"]);
                var configId = Convert.ToInt32(_decorOrdersDt.Rows[i]["decorConfigId"]);
                var measureId = Convert.ToInt32(_decorOrdersDt.Rows[i]["measureId"]);
                var length = Convert.ToInt32(_decorOrdersDt.Rows[i]["length"]);
                var height = Convert.ToInt32(_decorOrdersDt.Rows[i]["height"]);
                var width = Convert.ToInt32(_decorOrdersDt.Rows[i]["width"]);
                var count = Convert.ToInt32(_decorOrdersDt.Rows[i]["count"]);
                var countRate1 = 0;
                var decorName = TechStorename(decorId);

                if (_decorOrdersDt.Rows[i]["OrderDate"] != DBNull.Value)
                    OrderDate = Convert.ToDateTime(_decorOrdersDt.Rows[i]["OrderDate"]).ToShortDateString();

                //var rows1 = _machineRatesDt
                //    .AsEnumerable()
                //    .Where(row => Convert.ToInt32(row["configId"]) == configId);
                var rows = _machineRatesDt.Select($"configId={configId} and itemId={itemId}");

                if (!rows.Any())
                    continue;

                foreach (var row in rows)
                {
                    var userId = 0;
                    var machineId = Convert.ToInt32(row["machineId"]);
                    var sectorId = Convert.ToInt32(row["sectorId"]);
                    var rate1 = Convert.ToDecimal(row["rate1"]);
                    var rate2 = Convert.ToDecimal(row["rate2"]);
                    var rate3 = Convert.ToDecimal(row["rate3"]);
                    var rate4 = Convert.ToDecimal(row["rate4"]);

                    var sectorIndex = SectorsList.FindIndex(s => s.Id == sectorId);
                    if (sectorIndex == -1)
                        continue;

                    if (decorName.Length > 0)
                    {

                        var index11 = decorName.IndexOf(" ");

                        if (index11 == -1)
                        {
                            var rows1 = _distDecorDt.Select($@"decorName='{decorName}' and sectorId={sectorId} 
and mainOrderId={mainOrderId}");

                            if (!rows1.Any())
                            {
                                var newRow = _distDecorDt.NewRow();
                                newRow["decorName"] = decorName;
                                newRow["sectorId"] = sectorId;
                                newRow["mainOrderId"] = mainOrderId;
                                _distDecorDt.Rows.Add(newRow);
                                //SectorsList[sectorIndex].Total1 += rate1;
                                //SectorsList[sectorIndex].Total2 += rate2;
                                //SectorsList[sectorIndex].Total3 += rate3;
                                //SectorsList[sectorIndex].Total4 += rate4;
                                countRate1 = 1;
                            }
                        }
                        else
                        {
                            var s11 = decorName.Substring(0, index11);

                            var rows1 = _distDecorDt.Select(
                                $@"decorName='{s11}' and sectorId={sectorId} and mainOrderId={mainOrderId}");

                            if (!rows1.Any())
                            {
                                var newRow = _distDecorDt.NewRow();
                                newRow["decorName"] = s11;
                                newRow["sectorId"] = sectorId;
                                newRow["mainOrderId"] = mainOrderId;
                                _distDecorDt.Rows.Add(newRow);
                                //SectorsList[sectorIndex].Total1 += rate1;
                                //SectorsList[sectorIndex].Total2 += rate2;
                                //SectorsList[sectorIndex].Total3 += rate3;
                                //SectorsList[sectorIndex].Total4 += rate4;
                                countRate1 = 1;
                            }
                        }
                    }

                    AddClientOrder(clientId, ClientName(clientId), orderNumber, (int)orderStatus, OrderDate);

                    AddCalculationDecorRow(userId, machineId, sectorId, clientId, orderNumber, orderStatus,
                        decorId, measureId, length, height, width, count, countRate1,
                        rate1, rate2, rate3, rate4);
                }
            }

            for (var i = 0; i < _frontsOrdersDt.Rows.Count; i++)
            {
                var mainOrderId = Convert.ToInt32(_frontsOrdersDt.Rows[i]["mainOrderId"]);
                var clientId = Convert.ToInt32(_frontsOrdersDt.Rows[i]["clientId"]);
                var OrderDate = string.Empty;
                var orderNumber = Convert.ToInt32(_frontsOrdersDt.Rows[i]["orderNumber"]);
                var itemId = Convert.ToInt32(_frontsOrdersDt.Rows[i]["frontId"]);
                var configId = Convert.ToInt32(_frontsOrdersDt.Rows[i]["frontConfigId"]);
                var measureId = Convert.ToInt32(_frontsOrdersDt.Rows[i]["measureId"]);
                var height = Convert.ToInt32(_frontsOrdersDt.Rows[i]["height"]);
                var width = Convert.ToInt32(_frontsOrdersDt.Rows[i]["width"]);
                var count = Convert.ToInt32(_frontsOrdersDt.Rows[i]["count"]);
                var countRate1 = 0;
                var techStorename = TechStorename(itemId);

                if (_frontsOrdersDt.Rows[i]["OrderDate"] != DBNull.Value)
                    OrderDate = Convert.ToDateTime(_frontsOrdersDt.Rows[i]["OrderDate"]).ToShortDateString();

                //var rows1 = _machineRatesDt
                //    .AsEnumerable()
                //    .Where(row => Convert.ToInt32(row["configId"]) == configId);
                var rows = _machineRatesDt.Select($"configId={configId} and itemId={itemId}");

                if (!rows.Any())
                    continue;

                foreach (var row in rows)
                {
                    var userId = 0;
                    var machineId = Convert.ToInt32(row["machineId"]);
                    var sectorId = Convert.ToInt32(row["sectorId"]);
                    var rate1 = Convert.ToDecimal(row["rate1"]);
                    var rate2 = Convert.ToDecimal(row["rate2"]);
                    var rate3 = Convert.ToDecimal(row["rate3"]);
                    var rate4 = Convert.ToDecimal(row["rate4"]);

                    var sectorIndex = SectorsList.FindIndex(s => s.Id == sectorId);
                    if (sectorIndex == -1)
                        continue;

                    if (techStorename.Length > 0)
                    {

                        var index11 = techStorename.IndexOf(" ");

                        if (index11 == -1)
                        {
                            var rows1 = _distDecorDt.Select($@"decorName='{techStorename}' and sectorId={sectorId} 
and mainOrderId={mainOrderId}");

                            if (!rows1.Any())
                            {
                                var newRow = _distDecorDt.NewRow();
                                newRow["decorName"] = techStorename;
                                newRow["sectorId"] = sectorId;
                                newRow["mainOrderId"] = mainOrderId;
                                _distDecorDt.Rows.Add(newRow);
                                //SectorsList[sectorIndex].Total1 += rate1;
                                //SectorsList[sectorIndex].Total2 += rate2;
                                //SectorsList[sectorIndex].Total3 += rate3;
                                //SectorsList[sectorIndex].Total4 += rate4;
                                countRate1 = 1;
                            }
                        }
                        else
                        {
                            var s11 = techStorename.Substring(0, index11);

                            var rows1 = _distDecorDt.Select(
                                $@"decorName='{s11}' and sectorId={sectorId} and mainOrderId={mainOrderId}");

                            if (!rows1.Any())
                            {
                                var newRow = _distDecorDt.NewRow();
                                newRow["decorName"] = s11;
                                newRow["sectorId"] = sectorId;
                                newRow["mainOrderId"] = mainOrderId;
                                _distDecorDt.Rows.Add(newRow);
                                //SectorsList[sectorIndex].Total1 += rate1;
                                //SectorsList[sectorIndex].Total2 += rate2;
                                //SectorsList[sectorIndex].Total3 += rate3;
                                //SectorsList[sectorIndex].Total4 += rate4;
                                countRate1 = 1;
                            }
                        }
                    }

                    AddClientOrder(clientId, ClientName(clientId), orderNumber, (int)orderStatus, OrderDate);

                    AddCalculationFrontRow(userId, machineId, sectorId, clientId, orderNumber, orderStatus,
                        itemId, measureId, 0, height, width, count, countRate1,
                        rate1, rate2, rate3, rate4);
                }
            }
        }

        private void AddCalculationFrontRow(int userId, int machineId, int sectorId, int clientId, int orderNumber, OrderStatus orderStatus,
            int decorId, int measureId, decimal length, decimal height, decimal width, decimal count, decimal countRate1,
            decimal rate1, decimal rate2, decimal rate3, decimal rate4)
        {
            decimal measureCount = 0;

            switch (measureId)
            {
                case 1:
                    measureCount = decimal.Round(height * width / 1000000, 3, MidpointRounding.AwayFromZero) * count;
                    break;
                case 2:
                    if (height != -1)
                        measureCount = decimal.Round(height / 1000, 3, MidpointRounding.AwayFromZero) * count;
                    if (length != -1)
                        measureCount = decimal.Round(length / 1000, 3, MidpointRounding.AwayFromZero) * count;
                    break;
                case 3:
                    measureCount = count;
                    break;
            }

            var total1 = Math.Ceiling(rate1 * countRate1 / 0.0000001m) * 0.0000001m;
            var total2 = Math.Ceiling(rate2 * measureCount / 0.0000001m) * 0.0000001m;
            var total3 = Math.Ceiling(rate3 * measureCount / 0.0000001m) * 0.0000001m;
            var total4 = Math.Ceiling(rate4 * measureCount / 0.0000001m) * 0.0000001m;
            var sumTotal = total1 + total2 + total3 + total4;

            switch (orderStatus)
            {
                case OrderStatus.NotConfirmed:
                    NotConfirmedSumRank += sumTotal;
                    break;
                case OrderStatus.ForAgreed:
                    ForAgreedSumRank += sumTotal;
                    break;
                case OrderStatus.Agreed:
                    AgreedSumRank += sumTotal;
                    break;
                case OrderStatus.OnProduction:
                    OnProductionSumRank += sumTotal;
                    break;
                case OrderStatus.InProduction:
                    InProductionSumRank += sumTotal;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(orderStatus), orderStatus, null);
            }

            var newRow = _loadCalculationsDt.NewRow();
            newRow["count"] = count;
            newRow["userId"] = userId;
            newRow["machineId"] = machineId;
            newRow["sectorId"] = sectorId;
            newRow["clientId"] = clientId;
            newRow["orderNumber"] = orderNumber;
            newRow["orderStatus"] = (int)orderStatus;
            newRow["machineName"] = MachineName(machineId);
            newRow["decorName"] = TechStorename(decorId);
            newRow["total1"] = total1;
            newRow["total2"] = total2;
            newRow["total3"] = total3;
            newRow["total4"] = total4;
            newRow["rate1"] = rate1;
            newRow["rate2"] = rate2;
            newRow["rate3"] = rate3;
            newRow["rate4"] = rate4;
            if (sumTotal != 0)
                newRow["sumTotal"] = sumTotal;
            _loadCalculationsDt.Rows.Add(newRow);
        }

        private void AddCalculationDecorRow(int userId, int machineId, int sectorId, int clientId, int orderNumber, OrderStatus orderStatus,
            int decorId, int measureId, decimal length, decimal height, decimal width, decimal count, decimal countRate1,
            decimal rate1, decimal rate2, decimal rate3, decimal rate4)
        {
            decimal measureCount = 0;

            switch (measureId)
            {
                case 1:
                    if (height == -1)
                        measureCount = decimal.Round(length * width / 1000000, 3, MidpointRounding.AwayFromZero) * count;
                    if (length == -1)
                        measureCount = decimal.Round(height * width / 1000000, 3, MidpointRounding.AwayFromZero) * count;
                    break;
                case 2:
                    if (height != -1)
                        measureCount = decimal.Round(height / 1000, 3, MidpointRounding.AwayFromZero) * count;
                    if (length != -1)
                        measureCount = decimal.Round(length / 1000, 3, MidpointRounding.AwayFromZero) * count;
                    break;
                case 3:
                    measureCount = count;
                    break;
            }

            var total1 = Math.Ceiling(rate1 * countRate1 / 0.0000001m) * 0.0000001m;
            var total2 = Math.Ceiling(rate2 * measureCount / 0.0000001m) * 0.0000001m;
            var total3 = Math.Ceiling(rate3 * measureCount / 0.0000001m) * 0.0000001m;
            var total4 = Math.Ceiling(rate4 * measureCount / 0.0000001m) * 0.0000001m;
            var sumTotal = total1 + total2 + total3 + total4;

            switch (orderStatus)
            {
                case OrderStatus.NotConfirmed:
                    NotConfirmedSumRank += sumTotal;
                    break;
                case OrderStatus.ForAgreed:
                    ForAgreedSumRank += sumTotal;
                    break;
                case OrderStatus.Agreed:
                    AgreedSumRank += sumTotal;
                    break;
                case OrderStatus.OnProduction:
                    OnProductionSumRank += sumTotal;
                    break;
                case OrderStatus.InProduction:
                    InProductionSumRank += sumTotal;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(orderStatus), orderStatus, null);
            }

            var newRow = _loadCalculationsDt.NewRow();
            newRow["count"] = count;
            newRow["userId"] = userId;
            newRow["machineId"] = machineId;
            newRow["sectorId"] = sectorId;
            newRow["clientId"] = clientId;
            newRow["orderNumber"] = orderNumber;
            newRow["orderStatus"] = (int)orderStatus;
            newRow["machineName"] = MachineName(machineId);
            newRow["decorName"] = TechStorename(decorId);
            newRow["total1"] = total1;
            newRow["total2"] = total2;
            newRow["total3"] = total3;
            newRow["total4"] = total4;
            newRow["rate1"] = rate1;
            newRow["rate2"] = rate2;
            newRow["rate3"] = rate3;
            newRow["rate4"] = rate4;
            if (sumTotal != 0)
                newRow["sumTotal"] = sumTotal;
            _loadCalculationsDt.Rows.Add(newRow);
        }

        private void AddClientRow(int machineId, int sectorId, int clientId, int orderNumber, int orderStatus,
            decimal total1, decimal total2, decimal total3, decimal total4)
        {
            var newRow = _groupbyClientDt.NewRow();
            newRow["machineId"] = machineId;
            newRow["sectorId"] = sectorId;
            newRow["clientId"] = clientId;
            newRow["orderNumber"] = orderNumber;
            newRow["orderStatus"] = orderStatus;
            newRow["total1"] = total1;
            newRow["total2"] = total2;
            newRow["total3"] = total3;
            newRow["total4"] = total4;
            _groupbyClientDt.Rows.Add(newRow);
        }

        private void EditClientRow(int machineId, int sectorId, int clientId, int orderNumber, int orderStatus,
            decimal total1, decimal total2, decimal total3, decimal total4)
        {
            var filter = $"clientId={clientId} and orderNumber={orderNumber} and orderStatus={orderStatus} and sectorId={sectorId}";
            var rows = _groupbyClientDt.Select(filter);
            if (!rows.Any())
                return;

            rows[0]["machineId"] = machineId;
            rows[0]["total1"] = Convert.ToDecimal(rows[0]["total1"]) + total1;
            rows[0]["total2"] = Convert.ToDecimal(rows[0]["total2"]) + total2;
            rows[0]["total3"] = Convert.ToDecimal(rows[0]["total3"]) + total3;
            rows[0]["total4"] = Convert.ToDecimal(rows[0]["total4"]) + total4;
        }

        private void AddGroupByMachineRow(int machineId, int sectorId, int clientId, int orderNumber,
            decimal total1, decimal total2, decimal total3, decimal total4)
        {
            var newRow = _groupbyMachinesDt.NewRow();
            newRow["machineId"] = machineId;
            newRow["sectorId"] = sectorId;
            newRow["clientId"] = clientId;
            newRow["orderNumber"] = orderNumber;
            newRow["total1"] = total1;
            newRow["total2"] = total2;
            newRow["total3"] = total3;
            newRow["total4"] = total4;
            _groupbyMachinesDt.Rows.Add(newRow);
        }

        public void GetDecorOrdersNotConfirmed(bool fronts, bool decor)
        {
            var mainOrderFilter = $@" and newmainorders.factoryid = 1 and newmainorders.profilproductionstatusid=1
                and newmainorders.profilproductionstatusid<>2 
                and newmainorders.profilstoragestatusid<>2 
                and newmainorders.profilexpeditionstatusid<>2
                and newmainorders.profildispatchstatusid<>2";

            var megaOrderFilter = " and AgreementStatusID=1";

            var selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.decororderid, d.mainorderid,
d.decorid, d.productId, d.decorconfigid, dc.measureid, d.length, d.height, d.width, d.count from newdecororders  as d
inner join infiniu2_catalog.dbo.decorconfig as dc on d.decorconfigid = dc.decorconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 order by d.mainorderid";

            _decorOrdersDt.Clear();
            if (decor)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_decorOrdersDt);

            selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.FrontsOrdersID, d.mainorderid,
d.frontid, d.frontconfigid, dc.measureid, d.height, d.width, d.count from newfrontsorders  as d
inner join infiniu2_catalog.dbo.frontsconfig as dc on d.frontconfigid = dc.frontconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 order by d.mainorderid";

            _frontsOrdersDt.Clear();
            if (fronts)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_frontsOrdersDt);
        }

        public void GetDecorOrdersForAgreed(bool fronts, bool decor)
        {
            var mainOrderFilter = $@" and newmainorders.factoryid = 1 and newmainorders.profilproductionstatusid=1
                and newmainorders.profilproductionstatusid<>2 
                and newmainorders.profilstoragestatusid<>2 
                and newmainorders.profilexpeditionstatusid<>2
                and newmainorders.profildispatchstatusid<>2";

            var megaOrderFilter = " and AgreementStatusID=3";

            var selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.decororderid, d.mainorderid,
d.decorid, d.productId, d.decorconfigid, dc.measureid, d.length, d.height, d.width, d.count from newdecororders  as d
inner join infiniu2_catalog.dbo.decorconfig as dc on d.decorconfigid = dc.decorconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 order by d.mainorderid";

            _decorOrdersDt.Clear();
            if (decor)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_decorOrdersDt);

            selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.FrontsOrdersID, d.mainorderid,
d.frontid, d.frontconfigid, dc.measureid, d.height, d.width, d.count from newfrontsorders  as d
inner join infiniu2_catalog.dbo.frontsconfig as dc on d.frontconfigid = dc.frontconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 order by d.mainorderid";

            _frontsOrdersDt.Clear();
            if (fronts)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_frontsOrdersDt);
        }

        public void GetDecorOrdersAgreed(bool fronts, bool decor)
        {
            var mainOrderFilter = $@" and newmainorders.factoryid = 1 and newmainorders.profilproductionstatusid=1
                and newmainorders.profilproductionstatusid<>2 
                and newmainorders.profilstoragestatusid<>2 
                and newmainorders.profilexpeditionstatusid<>2
                and newmainorders.profildispatchstatusid<>2";

            var megaOrderFilter = " and AgreementStatusID=2";

            var selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.decororderid, d.mainorderid,
d.decorid, d.productId, d.decorconfigid, dc.measureid, d.length, d.height, d.width, d.count from newdecororders  as d
inner join infiniu2_catalog.dbo.decorconfig as dc on d.decorconfigid = dc.decorconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 and (d.mainorderid not in (select mainorderid from packages) or
d.decororderid not in (select orderid from packagedetails where packageid in (select packageid from packages where ProductType=1 and packagestatusid > 0)))
order by d.mainorderid";

            _decorOrdersDt.Clear();
            if (decor)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_decorOrdersDt);

            selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.FrontsOrdersID, d.mainorderid,
d.frontid, d.frontconfigid, dc.measureid, d.height, d.width, d.count from newfrontsorders  as d
inner join infiniu2_catalog.dbo.frontsconfig as dc on d.frontconfigid = dc.frontconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 and (d.mainorderid not in (select mainorderid from packages) or
d.FrontsOrdersID not in (select orderid from packagedetails where packageid in (select packageid from packages where ProductType=0 and packagestatusid > 0)))
order by d.mainorderid";

            _frontsOrdersDt.Clear();
            if (fronts)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_frontsOrdersDt);
        }

        public void GetDecorOrdersOnProduction(bool fronts, bool decor)
        {
            var mainOrderFilter = $@" and newmainorders.factoryid = 1 and newmainorders.profilproductionstatusid=3
                and newmainorders.profilproductionstatusid<>2 
                and newmainorders.profilstoragestatusid<>2 
                and newmainorders.profilexpeditionstatusid<>2
                and newmainorders.profildispatchstatusid<>2";

            var megaOrderFilter = " and AgreementStatusID=2";

            var selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.decororderid, d.mainorderid,
d.decorid, d.productId, d.decorconfigid, dc.measureid, d.length, d.height, d.width, d.count from newdecororders  as d
inner join infiniu2_catalog.dbo.decorconfig as dc on d.decorconfigid = dc.decorconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 and (d.mainorderid not in (select mainorderid from packages) or
d.decororderid not in (select orderid from packagedetails where packageid in (select packageid from packages where ProductType=1 and packagestatusid > 0)))
order by d.mainorderid";

            _decorOrdersDt.Clear();
            if (decor)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_decorOrdersDt);

            selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.FrontsOrdersID, d.mainorderid,
d.frontid, d.frontconfigid, dc.measureid, d.height, d.width, d.count from newfrontsorders  as d
inner join infiniu2_catalog.dbo.frontsconfig as dc on d.frontconfigid = dc.frontconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 and (d.mainorderid not in (select mainorderid from packages) or
d.FrontsOrdersID not in (select orderid from packagedetails where packageid in (select packageid from packages where ProductType=0 and packagestatusid > 0)))
order by d.mainorderid";

            _frontsOrdersDt.Clear();
            if (fronts)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_frontsOrdersDt);
        }

        public void GetDecorOrdersInProduction(bool fronts, bool decor)
        {
            var mainOrderFilter = $@" and newmainorders.factoryid = 1 and newmainorders.profilproductionstatusid=2";

            var megaOrderFilter = " and AgreementStatusID=2";

            var selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.decororderid, d.mainorderid,
d.decorid, d.productId, d.decorconfigid, dc.measureid, d.length, d.height, d.width, d.count from newdecororders  as d
inner join infiniu2_catalog.dbo.decorconfig as dc on d.decorconfigid = dc.decorconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 and (d.mainorderid not in (select mainorderid from packages) or
d.decororderid not in (select orderid from packagedetails where packageid in (select packageid from packages where ProductType=1 and packagestatusid > 0)))
order by d.mainorderid";

            _decorOrdersDt.Clear();
            if (decor)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_decorOrdersDt);

            selectCommand = $@"select m.megaorderid, m.ordernumber, m.OrderDate, m.clientid, d.FrontsOrdersID, d.mainorderid,
d.frontid, d.frontconfigid, dc.measureid, d.height, d.width, d.count from newfrontsorders  as d
inner join infiniu2_catalog.dbo.frontsconfig as dc on d.frontconfigid = dc.frontconfigid
inner join newmainorders on d.mainorderid = newmainorders.mainorderid {mainOrderFilter}
inner join newMegaorders as m on newMainorders.megaorderid = m.megaorderid {megaOrderFilter}
where onStorage = 0 and (d.mainorderid not in (select mainorderid from packages) or
d.FrontsOrdersID not in (select orderid from packagedetails where packageid in (select packageid from packages where ProductType=0 and packagestatusid > 0)))
order by d.mainorderid";

            _frontsOrdersDt.Clear();
            if (fronts)
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    da.Fill(_frontsOrdersDt);
        }

        public List<Sector> GetAllSectors()
        {
            var sectorsList = new List<Sector>();

            const string selectCommand = @"select sectorid, sectorname from sectors order by sectorname";

            using (var dt = new DataTable())
            {
                using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString))
                    da.Fill(dt);

                foreach (DataRow item in dt.Rows)
                {
                    var sector = new Sector
                    {
                        Name = item["sectorname"].ToString(),
                        Id = Convert.ToInt32(item["sectorid"])
                    };

                    sectorsList.Add(sector);
                }
            }

            return sectorsList;
        }

        private void GetClients()
        {
            const string selectCommand = @"select clientId, clientName from clients";

            _clientsDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingReferenceConnectionString)) da.Fill(_clientsDt);
        }

        private void GetTechStore()
        {
            const string selectCommand = @"select techstoreid, techstorename from techstore";

            _techstoreDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString)) da.Fill(_techstoreDt);
        }

        private void GetMachines()
        {
            const string selectCommand = @"select machines.machineid, sectors.sectorid, 
sectors.sectorname, subsectors.subsectorname, machines.machinename from machines 
inner join subsectors on machines.subsectorid = subsectors.subsectorid 
inner join sectors on subsectors.sectorid = sectors.sectorid";

            _machinesDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString)) da.Fill(_machinesDt);
        }

        private void GetMachinePriority()
        {
            const string selectCommand =
                @"select id, producttype, itemid, machineid, priority from machineitemsinhercules";

            _machinePriorityDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString)) da.Fill(_machinePriorityDt);
        }

        private void GetMachineRates()
        {
            const string selectCommand =
                @"select s.sectorid, s.sectorname, t.techstorename, m.* from machineratesinhercules as m 
left outer join decorconfig as dc on m.configid = dc.decorconfigid 
left outer join techstore as t on dc.decorid = t.techstoreid
inner join machines on m.machineid = machines.machineid 
inner join subsectors on machines.subsectorid = subsectors.subsectorid 
inner join sectors as s on subsectors.sectorid = s.sectorid";

            _machineRatesDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString)) da.Fill(_machineRatesDt);
        }

        private void GetStaffing()
        {
            const string selectCommand =
                @"select s.departmentid as sectorid, departments.departmentname, s.positionid, p.position, 
s.userid, u.shortname, s.rank, s.rate
from stafflist as s inner join departments on s.departmentid = departments.departmentid 
inner join positions as p on s.positionid = p.positionid 
inner join infiniu2_users.dbo.users as u on s.userid = u.userid";

            _staffingDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.LightConnectionString)) da.Fill(_staffingDt);
        }

        public class Client
        {
            public Client(string name, int clientId, int orderNumber, int orderStatus)
            {
                Name = name;
                ClientId = clientId;
                OrderNumber = orderNumber;
                OrderStatus = orderStatus;
            }

            public string Name
            {
                get;
                set;
            }

            public int ClientId
            {
                get;
                set;
            }

            public int OrderNumber
            {
                get;
                set;
            }

            public int OrderStatus
            {
                get;
                set;
            }

        }

        public class Machine
        {
            public Machine()
            {
            }

            public Machine(int id, string name)
            {
                Id = id;
                Name = name;
            }

            public DataTable DataSource1
            {
                get;
                set;
            }
            public BindingSource DataSourceNotConfirmed
            {
                get;
                set;
            }
            public BindingSource DataSourceForAgreed
            {
                get;
                set;
            }

            public BindingSource DataSourceAgreed
            {
                get;
                set;
            }

            public BindingSource DataSourceOnProduction
            {
                get;
                set;
            }

            public BindingSource DataSourceInProduction
            {
                get;
                set;
            }

            public int Id
            {
                get;
                set;
            }

            public int ClientId
            {
                get;
                set;
            }

            public string Name
            {
                get;
                set;
            }

            public decimal SumTotal
            {
                get;
                set;
            }
        }

        public class Sector
        {
            public Sector()
            {
            }

            public Sector(int id, string name)
            {
                Id = id;
                Name = name;
            }

            public int Id
            {
                get;
                set;
            }

            public decimal SumTotal
            {
                get;
                set;
            }

            public decimal Total1
            {
                get;
                set;
            }

            public decimal Total2
            {
                get;
                set;
            }

            public decimal Total3
            {
                get;
                set;
            }

            public decimal Total4
            {
                get;
                set;
            }

            public decimal Percent
            {
                get;
                set;
            }

            public decimal SumRank
            {
                get;
                set;
            }

            public decimal Rank1
            {
                get;
                set;
            }

            public decimal Rank2
            {
                get;
                set;
            }

            public decimal Rank3
            {
                get;
                set;
            }

            public decimal Rank4
            {
                get;
                set;
            }

            public string Name
            {
                get;
                set;
            }
        }
    }
}
