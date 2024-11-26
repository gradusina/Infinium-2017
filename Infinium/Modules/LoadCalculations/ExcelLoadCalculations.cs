using NPOI.HSSF.UserModel;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;


namespace Infinium.Modules.LoadCalculations
{

    public class Employee
    {
        public int ID { get; set; }
        public int configId { get; set; }
        public int itemId { get; set; }
        public decimal sumTotal { get; set; }
        public string accountName { get; set; }
        public string invnumber { get; set; }
        public decimal count { get; set; }
        public List<LoadCalculations.Sector> sectors { get; set; }
    }

    public class ExcelLoadCalculations
    {
        private readonly DataTable _decorConfigsDt;
        private readonly DataTable _frontsConfigsDt;

        private readonly DataTable _sectorsDt;
        private DataTable _ordersDt;
        private readonly DataTable _frontsOrdersDt;

        private readonly DataTable _groupbySectorsDt;
        private readonly DataTable _loadCalculationsDt;
        private readonly DataTable _machinePriorityDt;
        private readonly DataTable _machineRatesDt;
        private readonly DataTable _machinesDt;

        private readonly DataTable _staffingDt;
        private readonly DataTable _excelDt;

        private XSSFWorkbook wBook;
        List<Employee> listEmployees;

        public ExcelLoadCalculations()
        {
            listEmployees = new List<Employee>();

            _frontsConfigsDt = new DataTable();
            _decorConfigsDt = new DataTable();

            _frontsOrdersDt = new DataTable();
            _ordersDt = new DataTable();
            _sectorsDt = new DataTable();
            _machinePriorityDt = new DataTable();
            _machineRatesDt = new DataTable();
            _machinesDt = new DataTable();
            _staffingDt = new DataTable();

            _excelDt = new DataTable();
            _excelDt.Columns.Add("accountName", typeof(string));
            _excelDt.Columns.Add("invNumber", typeof(string));
            _excelDt.Columns.Add("count", typeof(decimal));
            _excelDt.Columns.Add("cost", typeof(decimal));
            _excelDt.TableName = "excelTable";

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
            _loadCalculationsDt.Columns.Add("sectorId", typeof(int));
            _loadCalculationsDt.Columns.Add("count", typeof(decimal));
            _loadCalculationsDt.Columns.Add("accountName", typeof(string));
            _loadCalculationsDt.Columns.Add("rate1", typeof(decimal));
            _loadCalculationsDt.Columns.Add("rate2", typeof(decimal));
            _loadCalculationsDt.Columns.Add("rate3", typeof(decimal));
            _loadCalculationsDt.Columns.Add("rate4", typeof(decimal));
            _loadCalculationsDt.Columns.Add("total1", typeof(decimal));
            _loadCalculationsDt.Columns.Add("total2", typeof(decimal));
            _loadCalculationsDt.Columns.Add("total3", typeof(decimal));
            _loadCalculationsDt.Columns.Add("total4", typeof(decimal));
            _loadCalculationsDt.Columns.Add("sumtotal", typeof(decimal));

            var selectCommand =
                $@"select frontconfigid, frontid, enabled, accountingname, invnumber from frontsconfig";

            using var da1 = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString);
            da1.Fill(_frontsConfigsDt);

            selectCommand =
                $@"select decorconfigid, productid, enabled, accountingname, invnumber from decorconfig";

            using var da2 = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString);
            da2.Fill(_decorConfigsDt);

            GetMachines();
            GetMachinePriority();
            GetMachineRates();
            GetStaffing();
        }

        public int RankCoef { get; set; } = 1;

        public List<DataTable> OrdersTables { get; set; } = new();
        public List<LoadCalculations.Sector> SectorsList { get; set; } = new();

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

        private string SectorName(int id)
        {
            var name = "";

            var rows = _machinesDt.Select($"sectorId={id}");
            if (rows.Any())
                name = rows[0]["sectorName"].ToString();
            return name;
        }

        private string SectorShortName(int id)
        {
            var name = "";

            var rows = _machinesDt.Select($"sectorId={id}");
            if (rows.Any())
                name = rows[0]["shortName"].ToString();
            return name;
        }

        public void CreateSectors()
        {
            SectorsList.Clear();

            SectorsList = GetAllSectors();
            //var sectorsId = new List<int>();

            //foreach (var sector in SectorsList)
            //{
            //    sectorsId.Add(sector.Id);
            //}
            //foreach (var sectorId in sectorsId)
            //{
            //    decimal rank1;
            //    decimal rank2;
            //    decimal rank3;
            //    decimal rank4;
            //    var sector = new Sector
            //    {
            //        Name = SectorName(sectorId),
            //        ShortName = SectorShortName(sectorId),
            //        Id = sectorId,
            //        Rank1 = rank1 = SumRatesByRank1(sectorId, 0) * RankCoef,
            //        Rank2 = rank2 = SumRatesByRank1(sectorId, 2) * RankCoef,
            //        Rank3 = rank3 = SumRatesByRank1(sectorId, 3) * RankCoef,
            //        Rank4 = rank4 = SumRatesByRank1(sectorId, 4) * RankCoef,

            //        SumRank = rank1 + rank2 + rank3 + rank4
            //    };
            //    SectorsList.Add(sector);
            //}

            //SectorsList = SectorsList.OrderByDescending(r => r.Name).ToList();
        }

        public void ClearCalculations()
        {
            _frontsOrdersDt.Clear();
            _ordersDt.Clear();
            _loadCalculationsDt.Clear();
        }

        //public void GetOrdersFromExcel()
        //{
        //    _excelDt.Clear();

        //    var s = Clipboard.GetText();
        //    var lines = s.Split('\n');
        //    var data = new List<string>(lines);

        //    if (data.Count > 0 && string.IsNullOrWhiteSpace(data[data.Count - 1]))
        //    {
        //        data.RemoveAt(data.Count - 1);
        //    }

        //    int index = 0;

        //    foreach (var iterationRow in data)
        //    {
        //        var row = iterationRow;
        //        if (row.EndsWith("\r"))
        //        {
        //            row = row.Substring(0, row.Length - "\r".Length);
        //        }

        //        var rowData = row.Split(new char[] { '\r', '\x09' });
        //        var newRow = _excelDt.NewRow();

        //        for (var i = 0; i < rowData.Length; i++)
        //        {
        //            if (i >= _excelDt.Columns.Count) break;

        //            if (rowData[i].Length > 0)
        //            {
        //                //if (_excelDt.Columns[i].ColumnName == "invNumber")
        //                //    rowData[i] = rowData[i].Replace(" ", "");
        //                //if (_excelDt.Columns[i].ColumnName == "accountName")
        //                //{
        //                //    if (accountNames.Any(rowData[i].Contains))
        //                //    {
        //                //        var table = _ordersDt.Clone();

        //                //        while (!accountNames.Any(rowData[i + 1].Contains))
        //                //        {
        //                //            var newRow1 = table.NewRow();
        //                //            newRow1[i] = rowData[i];
        //                //            table.Rows.Add(newRow1);
        //                //            i++;
        //                //        }
        //                //    }
        //                //}

        //                if (_excelDt.Columns[i].ColumnName == "accountName" && rowData[i] == "ТМЦ")
        //                    break;

        //                if (_excelDt.Columns[i].ColumnName == "invNumber")
        //                    newRow[i] = rowData[i].Replace(" ", "");
        //                else
        //                    newRow[i] = rowData[i];
        //            }
        //        }
        //        _excelDt.Rows.Add(newRow);
        //        index++;
        //    }
        //}

        //public void GetOrders()
        //{
        //    var rowIndex = 0;
        //    foreach (var productType in productTypes)
        //    {
        //        var toTable = false;
        //        var producttype = 0;
        //        var table = _ordersDt.Clone();

        //        for (var i = rowIndex; i < _excelDt.Rows.Count; i++)
        //        {
        //            var accountName = _excelDt.Rows[i]["accountName"].ToString();
        //            var invNumber = _excelDt.Rows[i]["invNumber"].ToString();
        //            decimal.TryParse(invNumber, out var count);

        //            if (productType == accountName)
        //            {
        //                toTable = true;
        //                producttype++;
        //                continue;
        //            }

        //            if (productTypes.Any(accountName.Contains))
        //            {
        //                rowIndex = i;
        //                break;
        //            }

        //            if (!toTable)
        //                continue;

        //            var configId = -1;
        //            var itemId = -1;

        //            var rows = _frontsConfigsDt.Select($"invNumber='{invNumber}'");
        //            if (rows.Any())
        //            {
        //                int.TryParse(rows[0]["frontconfigid"].ToString(), out configId);
        //                int.TryParse(rows[0]["frontid"].ToString(), out itemId);
        //            }
        //            else
        //            {

        //                rows = _decorConfigsDt.Select($"invNumber='{invNumber}'");
        //                if (rows.Any())
        //                {
        //                    int.TryParse(rows[0]["decorconfigid"].ToString(), out configId);
        //                    int.TryParse(rows[0]["decorid"].ToString(), out itemId);
        //                }
        //            }

        //            var newRow = table.NewRow();
        //            newRow["producttype"] = producttype;
        //            newRow["configId"] = configId;
        //            newRow["itemId"] = itemId;
        //            newRow["accountname"] = accountName;
        //            newRow["invNumber"] = invNumber;
        //            newRow["count"] = count;
        //            table.Rows.Add(newRow);

        //        }
        //        OrdersTables.Add(table);
        //    }
        //}

        //        public void ExportFile(string path)
        //        {
        //            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
        //{

        //                wBook = new XSSFWorkbook(fileStream);
        //            }

        //            listEmployees = new List<Employee>();
        //            var sheet = wBook.GetSheetAt(0);
        //            for (var i = 1; i <= sheet.LastRowNum - 1; i++)
        //            {
        //                if (sheet.GetRow(i) != null)
        //                {
        //                    var accountName = sheet.GetRow(i).GetCell(0).StringCellValue;
        //                    var invNumber = sheet.GetRow(i).GetCell(1).StringCellValue.Replace(" ", "");
        //                    var count = (decimal)sheet.GetRow(i).GetCell(2).NumericCellValue;
        //                    var cost = sheet.GetRow(i).GetCell(3).NumericCellValue;
        //                    var configId = -1;
        //                    var itemId = -1;

        //                    var rows = _frontsConfigsDt.Select($"invNumber='{invNumber}'");
        //                    if (rows.Any())
        //                    {
        //                        int.TryParse(rows[0]["frontconfigid"].ToString(), out configId);
        //                        int.TryParse(rows[0]["frontid"].ToString(), out itemId);
        //                    }
        //                    else
        //                    {

        //                        rows = _decorConfigsDt.Select($"invNumber='{invNumber}'");
        //                        if (rows.Any())
        //                        {
        //                            int.TryParse(rows[0]["decorconfigid"].ToString(), out configId);
        //                            int.TryParse(rows[0]["productid"].ToString(), out itemId);
        //                        }
        //                    }

        //                    var emp = new Employee
        //                    {
        //                        accountName = accountName,
        //                        invnumber = invNumber,
        //                        configId = configId,
        //                        itemId = itemId,
        //                        count = count
        //                    };
        //                    listEmployees.Add(emp);
        //                }
        //            }
        //        }

        //public void CalculateLoad(DataTable table)
        //{
        //    for (var i = 0; i < table.Rows.Count; i++)
        //    {
        //        var itemId = Convert.ToInt32(table.Rows[i]["itemId"]);
        //        var configId = Convert.ToInt32(table.Rows[i]["configId"]);
        //        var count = Convert.ToInt32(table.Rows[i]["count"]);
        //        var countRate1 = 0;
        //        var accountingName = table.Rows[i]["accountingname"].ToString();

        //        var rows = _machineRatesDt.Select($"configId={configId} and itemId={itemId}");

        //        if (!rows.Any())
        //            continue;

        //        foreach (var row in rows)
        //        {
        //            var sectorId = Convert.ToInt32(row["sectorId"]);
        //            var rate1 = Convert.ToDecimal(row["rate1"]);
        //            var rate2 = Convert.ToDecimal(row["rate2"]);
        //            var rate3 = Convert.ToDecimal(row["rate3"]);
        //            var rate4 = Convert.ToDecimal(row["rate4"]);

        //            var sectorIndex = SectorsList.FindIndex(s => s.Id == sectorId);
        //            if (sectorIndex == -1)
        //                continue;

        //            AddCalculationRow(sectorId, accountingName,
        //                count, countRate1,
        //                rate1, rate2, rate3, rate4);
        //        }
        //    }
        //}

        public void ExportFile(string path)
        {
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {

                wBook = new XSSFWorkbook(fileStream);
            }

            listEmployees = new List<Employee>();
            var sheet = wBook.GetSheetAt(0);
            for (var i = 1; i <= sheet.LastRowNum - 1; i++)
            {
                if (sheet.GetRow(i) != null)
                {
                    var accountName = sheet.GetRow(i).GetCell(0).StringCellValue;
                    var invNumber = sheet.GetRow(i).GetCell(1).StringCellValue.Replace(" ", "");
                    var count = (decimal)sheet.GetRow(i).GetCell(2).NumericCellValue;
                    var cost = sheet.GetRow(i).GetCell(3).NumericCellValue;
                    var configId = -1;
                    var itemId = -1;

                    var rows = _frontsConfigsDt.Select($"invNumber='{invNumber}'");
                    if (rows.Any())
                    {
                        int.TryParse(rows[0]["frontconfigid"].ToString(), out configId);
                        int.TryParse(rows[0]["frontid"].ToString(), out itemId);
                    }
                    else
                    {

                        rows = _decorConfigsDt.Select($"invNumber='{invNumber}'");
                        if (rows.Any())
                        {
                            int.TryParse(rows[0]["decorconfigid"].ToString(), out configId);
                            int.TryParse(rows[0]["productid"].ToString(), out itemId);
                        }
                    }

                    var emp = new Employee
                    {
                        accountName = accountName,
                        invnumber = invNumber,
                        configId = configId,
                        itemId = itemId,
                        count = count
                    };
                    listEmployees.Add(emp);
                }
            }
        }

        public void CalculateLoad()
        {
            var index = 0;
            decimal sumtotal1 = 0;
            decimal sumtotal2 = 0;
            decimal sumtotal3 = 0;
            decimal sumtotal4 = 0;
            decimal sumTotal = 0;

            for (var i = 0; i < listEmployees.Count; i++)
            {
                var configId = listEmployees[i].configId;
                var itemId = listEmployees[i].itemId;
                var rows = _machineRatesDt.Select($"configId={listEmployees[i].configId} and itemId={listEmployees[i].itemId}");

                listEmployees[i].sectors = NewSectorsList();

                if (listEmployees[i].invnumber.Length == 0)
                {
                    index = i;
                }

                if (!rows.Any())
                    continue;
                sumTotal = 0;
                sumtotal1 = 0;
                sumtotal2 = 0;
                sumtotal3 = 0;
                sumtotal4 = 0;

                foreach (var row in rows)
                {
                    var sectorId = Convert.ToInt32(row["sectorId"]);
                    var rate1 = Convert.ToDecimal(row["rate1"]);
                    var rate2 = Convert.ToDecimal(row["rate2"]);
                    var rate3 = Convert.ToDecimal(row["rate3"]);
                    var rate4 = Convert.ToDecimal(row["rate4"]);

                    var sectorIndex = SectorsList.FindIndex(s => s.Id == sectorId);
                    if (sectorIndex == -1)
                        continue;

                    var measureCount = listEmployees[i].count;
                    var total1 = Math.Ceiling(rate1 * 0 / 0.0000001m) * 0.0000001m;
                    var total2 = Math.Ceiling(rate2 * measureCount / 0.0000001m) * 0.0000001m;
                    var total3 = Math.Ceiling(rate3 * measureCount / 0.0000001m) * 0.0000001m;
                    var total4 = Math.Ceiling(rate4 * measureCount / 0.0000001m) * 0.0000001m;
                    sumtotal1 += total1;
                    sumtotal2 += total2;
                    sumtotal3 += total3;
                    sumtotal4 += total4;
                    sumTotal += total1 + total2 + total3 + total4;

                    listEmployees[i].sectors[sectorIndex].SumTotal += (total1 + total2 + total3 + total4);
                    listEmployees[index].sectors[sectorIndex].SumTotal += (total1 + total2 + total3 + total4);
                }
                listEmployees[i].sumTotal = sumTotal;
                listEmployees[index].sumTotal += sumTotal;
            }
        }

        public void CreateReportTable()
        {
            _ordersDt = new DataTable();
            _ordersDt.Columns.Add(new DataColumn("id", typeof(int)) { AutoIncrement = true });
            _ordersDt.Columns.Add("configid", typeof(int));
            _ordersDt.Columns.Add("itemid", typeof(int));
            _ordersDt.Columns.Add("accountName", typeof(string));
            _ordersDt.Columns.Add("invnumber", typeof(string));
            _ordersDt.Columns.Add("count", typeof(decimal));
            foreach (var sector in SectorsList)
            {
                _ordersDt.Columns.Add(sector.ShortName, typeof(decimal));
            }

            _ordersDt.Columns.Add("sumtotal", typeof(decimal));

            foreach (var item in listEmployees)
            {
                var newRow = _ordersDt.NewRow();
                newRow["configId"] = item.configId;
                newRow["itemId"] = item.itemId;
                if (item.sectors != null)
                    foreach (var sector in item.sectors)
                        newRow[sector.ShortName] = sector.SumTotal;
                newRow["accountName"] = item.accountName;
                newRow["invnumber"] = item.invnumber;
                newRow["count"] = item.count;
                newRow["sumTotal"] = item.sumTotal;
                _ordersDt.Rows.Add(newRow);
            }
        }

        public DataTable ReportDt => _ordersDt;
        private void AddCalculationRow(int sectorId,
            string accountingName,
            decimal count,
            decimal rate1, decimal rate2, decimal rate3, decimal rate4)
        {
            var measureCount = count;

            var total1 = Math.Ceiling(rate1 * 0 / 0.0000001m) * 0.0000001m;
            var total2 = Math.Ceiling(rate2 * measureCount / 0.0000001m) * 0.0000001m;
            var total3 = Math.Ceiling(rate3 * measureCount / 0.0000001m) * 0.0000001m;
            var total4 = Math.Ceiling(rate4 * measureCount / 0.0000001m) * 0.0000001m;
            var sumTotal = total1 + total2 + total3 + total4;


            var newRow = _loadCalculationsDt.NewRow();
            newRow["count"] = count;
            newRow["sectorId"] = sectorId;
            newRow["accountName"] = accountingName;
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

        public List<LoadCalculations.Sector> GetAllSectors()
        {
            var sectorsList = new List<LoadCalculations.Sector>();

            const string selectCommand = @"select sectorid, sectorname, shortName from sectors where sectorid<10051 order by sectorname";
            _sectorsDt.Clear();
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString))
            {
                da.Fill(_sectorsDt);
            }

            foreach (DataRow item in _sectorsDt.Rows)
            {
                var sector = new LoadCalculations.Sector
                {
                    Name = item["sectorname"].ToString(),
                    ShortName = item["shortName"].ToString(),
                    Id = Convert.ToInt32(item["sectorid"])
                };

                sectorsList.Add(sector);
            }

            return sectorsList;
        }

        public List<LoadCalculations.Sector> NewSectorsList()
        {
            var sectorsList = new List<LoadCalculations.Sector>();

            foreach (DataRow item in _sectorsDt.Rows)
            {
                var sector = new LoadCalculations.Sector
                {
                    Name = item["sectorname"].ToString(),
                    ShortName = item["shortName"].ToString(),
                    Id = Convert.ToInt32(item["sectorid"])
                };

                sectorsList.Add(sector);
            }

            return sectorsList;
        }

        private void GetMachines()
        {
            const string selectCommand = @"select machines.machineid, sectors.sectorid, 
sectors.sectorname, sectors.shortname, subsectors.subsectorname, machines.machinename from machines 
inner join subsectors on machines.subsectorid = subsectors.subsectorid 
inner join sectors on subsectors.sectorid = sectors.sectorid";

            _machinesDt.Clear();
            using var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString);
            da.Fill(_machinesDt);
        }

        private void GetMachinePriority()
        {
            const string selectCommand =
                @"select id, producttype, itemid, machineid, priority from machineitemsinhercules";

            _machinePriorityDt.Clear();
            using var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString);
            da.Fill(_machinePriorityDt);
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
            using var da = new SqlDataAdapter(selectCommand, ConnectionStrings.CatalogConnectionString);
            da.Fill(_machineRatesDt);
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
            using var da = new SqlDataAdapter(selectCommand, ConnectionStrings.LightConnectionString);
            da.Fill(_staffingDt);
        }
    }
}