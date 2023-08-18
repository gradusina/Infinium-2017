using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

using static Infinium.UserProfile;

namespace Infinium.Modules.ZOV.Clients
{
    public class Clients
    {
        private PercentageDataGrid _clientsDataGrid = null;
        private PercentageDataGrid _clientsGroupsDataGrid = null;
        private PercentageDataGrid _managersDataGrid = null;
        private PercentageDataGrid _contactsDataGrid = null;
        private PercentageDataGrid _shopAddressesDataGrid = null;

        private DataTable _clientsGroupsDataTable = null;
        public DataTable AllClientsDataTable = null;
        private DataTable _clientsDataTable = null;
        private DataTable _managersDataTable = null;

        public bool NewClient = false;

        public DataTable ContactsDataTable = null;
        public DataTable NewContactsDataTable = null;

        public DataTable ShopAddressesDataTable = null;
        public DataTable NewShopAddressesDataTable = null;

        public SqlDataAdapter ClientsGroupsDataAdapter = null;
        public SqlDataAdapter ManagersDataAdapter = null;
        public SqlDataAdapter ClientsDataAdapter = null;
        public SqlDataAdapter ShopAddressesDataAdapter = null;

        public SqlCommandBuilder ClientsGroupsCommandBuilder = null;
        public SqlCommandBuilder ManagersCommandBuilder = null;
        public SqlCommandBuilder ClientsCommandBuilder = null;
        public SqlCommandBuilder ShopAddressesCommandBuilder = null;

        public BindingSource ContactsBindingSource = null;
        public BindingSource NewContactsBindingSource = null;
        public BindingSource ManagersBindingSource = null;
        public BindingSource ClientsGroupsBindingSource = null;
        public BindingSource ClientsBindingSource = null;
        public BindingSource ShopAddressesBindingSource = null;
        public BindingSource NewShopAddressesBindingSource = null;

        public string ClientsBindingSourceDisplayMember = null;
        public string ClientsBindingSourceValueMember = null;


        public Clients(ref PercentageDataGrid tClientsDataGrid, ref PercentageDataGrid tClientsGroupsDataGrid,
            ref PercentageDataGrid tManagersDataGrid,
            ref PercentageDataGrid tContactsDataGrid, ref PercentageDataGrid tShopAddressesDataGrid)
        {
            _clientsDataGrid = tClientsDataGrid;
            _clientsGroupsDataGrid = tClientsGroupsDataGrid;
            _managersDataGrid = tManagersDataGrid;
            _contactsDataGrid = tContactsDataGrid;
            _shopAddressesDataGrid = tShopAddressesDataGrid;

            Initialize();
        }

        private void Create()
        {
            AllClientsDataTable = new DataTable();
            _clientsGroupsDataTable = new DataTable();
            _clientsDataTable = new DataTable();
            _managersDataTable = new DataTable();
            ShopAddressesDataTable = new DataTable();

            ContactsBindingSource = new BindingSource();
            NewContactsBindingSource = new BindingSource();
            ClientsGroupsBindingSource = new BindingSource();
            ClientsBindingSource = new BindingSource();
            ManagersBindingSource = new BindingSource();
            ShopAddressesBindingSource = new BindingSource();
            NewShopAddressesBindingSource = new BindingSource();
        }

        private void Fill()
        {
            ManagersDataAdapter = new SqlDataAdapter("SELECT ManagerID, Name FROM Managers ORDER BY Name",
                ConnectionStrings.ZOVReferenceConnectionString);
            ManagersCommandBuilder = new SqlCommandBuilder(ManagersDataAdapter);
            ManagersDataAdapter.Fill(_managersDataTable);

            ClientsGroupsDataAdapter = new SqlDataAdapter("SELECT * FROM ClientsGroups ORDER BY ClientGroupName",
                ConnectionStrings.ZOVReferenceConnectionString);
            ClientsGroupsCommandBuilder = new SqlCommandBuilder(ClientsGroupsDataAdapter);
            ClientsGroupsDataAdapter.Fill(_clientsGroupsDataTable);
            
            ClientsDataAdapter = new SqlDataAdapter("SELECT * FROM Clients ORDER BY ClientName",
                ConnectionStrings.ZOVReferenceConnectionString);
            ClientsCommandBuilder = new SqlCommandBuilder(ClientsDataAdapter);
            ClientsDataAdapter.Fill(_clientsDataTable);

            using (var da = new SqlDataAdapter(@"SELECT Clients.ClientName, ClientsGroups.ClientGroupName, Managers.Name
FROM Clients 
INNER JOIN Managers ON Clients.ManagerID = Managers.ManagerID 
INNER JOIN ClientsGroups ON Clients.ClientGroupID = ClientsGroups.ClientGroupID
ORDER BY Clients.ClientName", ConnectionStrings.ZOVReferenceConnectionString))
            {
                da.Fill(AllClientsDataTable);
            }

            ShopAddressesDataAdapter = new SqlDataAdapter("SELECT TOP 0 * FROM ShopAddresses ORDER BY Address",
                ConnectionStrings.ZOVReferenceConnectionString);
            ShopAddressesCommandBuilder = new SqlCommandBuilder(ShopAddressesDataAdapter);
            ShopAddressesDataAdapter.Fill(ShopAddressesDataTable);
        }

        private void Binding()
        {
            ClientsGroupsBindingSource.DataSource = _clientsGroupsDataTable;
            ClientsBindingSource.DataSource = _clientsDataTable;
            ManagersBindingSource.DataSource = _managersDataTable;
            ContactsBindingSource.DataSource = ContactsDataTable;
            ShopAddressesBindingSource.DataSource = ShopAddressesDataTable;

            _managersDataGrid.DataSource = ManagersBindingSource;
            _clientsGroupsDataGrid.DataSource = ClientsGroupsBindingSource;
            _clientsDataGrid.DataSource = ClientsBindingSource;
            _contactsDataGrid.DataSource = ContactsBindingSource;
            _shopAddressesDataGrid.DataSource = ShopAddressesBindingSource;

            ClientsBindingSourceDisplayMember = "ClientName";

            ClientsBindingSourceValueMember = "ClientID";
        }

        private void SetClientsDataGrid()
        {
            _clientsDataGrid.Columns.Add(ClientsGroupsColumn);
            _clientsDataGrid.Columns.Add(ManagerColumn);

            foreach (DataGridViewColumn column in _clientsDataGrid.Columns)
            {
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            _managersDataGrid.Columns["Name"].HeaderText = "Менеджер";

            _clientsGroupsDataGrid.Columns["ClientGroupName"].HeaderText = "Название";

            _clientsDataGrid.Columns["ClientGroupID"].Visible = false;
            _clientsDataGrid.Columns["ManagerID"].Visible = false;
            _clientsDataGrid.Columns["Contacts"].Visible = false;
            _clientsDataGrid.Columns["MoveOk"].Visible = false;

            _clientsDataGrid.Columns["ClientID"].HeaderText = "ID";
            _clientsDataGrid.Columns["ClientName"].HeaderText = "Название организации";

            _clientsDataGrid.AutoGenerateColumns = false;

            _clientsDataGrid.Columns["ClientID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            _clientsDataGrid.Columns["ClientID"].Width = 50;
            //ClientsDataGrid.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //ClientsDataGrid.Columns["ClientName"].MinimumWidth = 150;
            //ClientsDataGrid.Columns["ManagerColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //ClientsDataGrid.Columns["ManagerColumn"].MinimumWidth = 150;
            //ClientsDataGrid.Columns["ClientsGroupsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //ClientsDataGrid.Columns["ClientsGroupsColumn"].MinimumWidth = 150;

            var displayIndex = 0;
            _clientsDataGrid.Columns["ClientID"].DisplayIndex = displayIndex++;
            _clientsDataGrid.Columns["ClientName"].DisplayIndex = displayIndex++;
            _clientsDataGrid.Columns["ClientsGroupsColumn"].DisplayIndex = displayIndex++;
            _clientsDataGrid.Columns["ManagerColumn"].DisplayIndex = displayIndex++;
        }

        private DataGridViewComboBoxColumn ManagerColumn
        {
            get
            {
                var column = new DataGridViewComboBoxColumn()
                {
                    Name = "ManagerColumn",
                    HeaderText = "Менеджер",
                    DataPropertyName = "ManagerID",

                    DataSource = new DataView(_managersDataTable),
                    ValueMember = "ManagerID",
                    DisplayMember = "Name",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    ReadOnly = false
                };
                return column;
            }
        }

        private DataGridViewComboBoxColumn ClientsGroupsColumn
        {
            get
            {
                var column = new DataGridViewComboBoxColumn()
                {
                    Name = "ClientsGroupsColumn",
                    HeaderText = "Группа клиентов",
                    DataPropertyName = "ClientGroupID",

                    DataSource = new DataView(_clientsGroupsDataTable),
                    ValueMember = "ClientGroupID",
                    DisplayMember = "ClientGroupName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    ReadOnly = false
                };
                return column;
            }
        }

        private void SetContactsDataGrid()
        {
            _contactsDataGrid.Columns["Name"].HeaderText = "Имя";
            _contactsDataGrid.Columns["Position"].HeaderText = "Должность";
            _contactsDataGrid.Columns["Phone"].HeaderText = "Телефон";
            _contactsDataGrid.Columns["Email"].HeaderText = "E-mail";
            _contactsDataGrid.Columns["ICQ"].HeaderText = "ICQ";
            _contactsDataGrid.Columns["Notes"].HeaderText = "Примечание";

            _contactsDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            _shopAddressesDataGrid.Columns["Lat"].Visible = false;
            _shopAddressesDataGrid.Columns["Long"].Visible = false;
            _shopAddressesDataGrid.Columns["ShopAddressID"].Visible = false;
            _shopAddressesDataGrid.Columns["ClientID"].Visible = false;
            _shopAddressesDataGrid.Columns["City"].HeaderText = "Город";
            _shopAddressesDataGrid.Columns["Country"].HeaderText = "Страна";
            _shopAddressesDataGrid.Columns["Phone"].HeaderText = "Телефон";
            _shopAddressesDataGrid.Columns["Address"].HeaderText = "Адрес салона";
            _shopAddressesDataGrid.Columns["IsFronts"].HeaderText = "Фасады";
            _shopAddressesDataGrid.Columns["IsProfile"].HeaderText = "Погонаж";
            _shopAddressesDataGrid.Columns["IsFurniture"].HeaderText = "Фурнитура";

            _shopAddressesDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            _shopAddressesDataGrid.Columns["Address"].DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;
            _shopAddressesDataGrid.Columns["Address"].DefaultCellStyle.Font = new System.Drawing.Font("SEGOE UI", 13.0F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Pixel, ((byte)(0)));
        }

        private void Initialize()
        {
            Create();
            Fill();
            CreateContactsDataTable();
            Binding();
            SetClientsDataGrid();
            SetContactsDataGrid();
        }

        private string GetContactsXml(DataTable dt)
        {
            var sw = new StringWriter();
            dt.WriteXml(sw);

            return sw.ToString();
        }

        public void FillContactsDataTable(int clientId)
        {
            ContactsDataTable.Clear();

            var rows = _clientsDataTable.Select("ClientID = " + clientId);

            if (rows[0]["Contacts"].ToString().Length == 0)
                return;

            var contactXml = rows[0]["Contacts"].ToString();

            using (var sr = new StringReader(contactXml))
            {
                ContactsDataTable.ReadXml(sr);
            }
        }

        public void FillShopAddressesDataTable(int clientId)
        {
            ShopAddressesDataAdapter.Dispose();
            ShopAddressesCommandBuilder.Dispose();
            ShopAddressesDataAdapter = new SqlDataAdapter("SELECT * FROM ShopAddresses WHERE ClientID=" + clientId + " ORDER BY Address",
                ConnectionStrings.ZOVReferenceConnectionString);
            ShopAddressesCommandBuilder = new SqlCommandBuilder(ShopAddressesDataAdapter);
            ShopAddressesDataTable.Clear();
            ShopAddressesDataAdapter.Fill(ShopAddressesDataTable);
        }

        private void CreateContactsDataTable()
        {
            ContactsDataTable = new DataTable()
            {
                TableName = "ContactsDataTable"
            };
            ContactsDataTable.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            ContactsDataTable.Columns.Add(new DataColumn("Position", Type.GetType("System.String")));
            ContactsDataTable.Columns.Add(new DataColumn("Phone", Type.GetType("System.String")));
            ContactsDataTable.Columns.Add(new DataColumn("Email", Type.GetType("System.String")));
            ContactsDataTable.Columns.Add(new DataColumn("ICQ", Type.GetType("System.String")));
            ContactsDataTable.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
        }

        public void CreateNewContactsDataTable(ref PercentageDataGrid dataGrid, ref PercentageDataGrid tShopAddressesDataGrid)
        {
            if (NewContactsDataTable == null)
            {
                NewContactsDataTable = new DataTable();
                NewContactsDataTable = ContactsDataTable.Clone();
            }
            else
                NewContactsDataTable.Clear();

            NewContactsBindingSource.DataSource = NewContactsDataTable;
            dataGrid.DataSource = NewContactsBindingSource;

            dataGrid.Columns["Name"].HeaderText = "Имя";
            dataGrid.Columns["Position"].HeaderText = "Должность";
            dataGrid.Columns["Phone"].HeaderText = "Телефон";
            dataGrid.Columns["Email"].HeaderText = "E-mail";
            dataGrid.Columns["ICQ"].HeaderText = "ICQ";
            dataGrid.Columns["Notes"].HeaderText = "Примечание";

            dataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            if (NewShopAddressesDataTable == null)
            {
                NewShopAddressesDataTable = new DataTable();
                NewShopAddressesDataTable = ShopAddressesDataTable.Clone();
            }
            else
                NewShopAddressesDataTable.Clear();
            NewShopAddressesBindingSource.DataSource = NewShopAddressesDataTable;
            tShopAddressesDataGrid.DataSource = NewShopAddressesBindingSource;

            tShopAddressesDataGrid.Columns["Lat"].Visible = false;
            tShopAddressesDataGrid.Columns["Long"].Visible = false;
            tShopAddressesDataGrid.Columns["ShopAddressID"].Visible = false;
            tShopAddressesDataGrid.Columns["ClientID"].Visible = false;
            tShopAddressesDataGrid.Columns["City"].HeaderText = "Город";
            tShopAddressesDataGrid.Columns["Country"].HeaderText = "Страна";
            tShopAddressesDataGrid.Columns["Phone"].HeaderText = "Телефон";
            tShopAddressesDataGrid.Columns["Address"].HeaderText = "Адрес салона";
            tShopAddressesDataGrid.Columns["IsFronts"].HeaderText = "Фасады";
            tShopAddressesDataGrid.Columns["IsProfile"].HeaderText = "Погонаж";
            tShopAddressesDataGrid.Columns["IsFurniture"].HeaderText = "Фурнитура";

            tShopAddressesDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            //tShopAddressesDataGrid.Columns["Address"].DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;
            //tShopAddressesDataGrid.Columns["Address"].DefaultCellStyle.Font = new System.Drawing.Font("SEGOE UI", 13.0F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Pixel, ((byte)(0)));
        }

        public void SaveClientsGroups()
        {
            ClientsGroupsDataAdapter.Update(_clientsGroupsDataTable);
            _clientsGroupsDataTable.Clear();
            ClientsGroupsDataAdapter.Fill(_clientsGroupsDataTable);
        }

        public void SaveAllClients()
        {
            ClientsDataAdapter.Update(_clientsDataTable);
            _clientsDataTable.Clear();
            ClientsDataAdapter.Fill(_clientsDataTable);
        }

        public void SaveManagers()
        {
            ManagersDataAdapter.Update(_managersDataTable);
            _managersDataTable.Clear();
            ManagersDataAdapter.Fill(_managersDataTable);
        }

        public void AddClient(string name, int clientGroupId, int managerId)
        {
            var contacts = GetContactsXml(NewContactsDataTable);

            var row = _clientsDataTable.NewRow();
            row["ManagerID"] = managerId;
            row["ClientName"] = name;
            row["ClientGroupID"] = clientGroupId;
            row["Contacts"] = contacts;

            _clientsDataTable.Rows.Add(row);
            ClientsDataAdapter.Update(_clientsDataTable);
            _clientsDataTable.Clear();
            ClientsDataAdapter.Fill(_clientsDataTable);
        }

        public void EditClient(ref string clientsName, ref int clientGroupId, ref int managerId)
        {
            clientsName = ((DataRowView)ClientsBindingSource.Current)["ClientName"].ToString();
            clientGroupId = Convert.ToInt32(((DataRowView)ClientsBindingSource.Current)["ClientGroupID"]);
            managerId = Convert.ToInt32(((DataRowView)ClientsBindingSource.Current)["ManagerID"]);

            NewContactsDataTable.Clear();
            NewContactsDataTable = ContactsDataTable.Copy();
            NewContactsBindingSource.DataSource = NewContactsDataTable;

            NewShopAddressesDataTable.Clear();
            NewShopAddressesDataTable = ShopAddressesDataTable.Copy();
            NewShopAddressesBindingSource.DataSource = NewShopAddressesDataTable;
        }

        public void SaveClient(string name, int clientGroupId, int managerId, int clientId)
        {
            var contacts = GetContactsXml(NewContactsDataTable);

            var index = ClientsBindingSource.Position;

            using (var da = new SqlDataAdapter(@"SELECT * FROM Clients WHERE ClientID = " + clientId + " ORDER BY ClientName",
                ConnectionStrings.ZOVReferenceConnectionString))
            {
                using (var cb = new SqlCommandBuilder(da))
                {
                    using (var dt = new DataTable())
                    {
                        da.Fill(dt);
                        if (dt.Rows.Count > 0)
                        {
                            dt.Rows[0]["ClientName"] = name;
                            dt.Rows[0]["ClientGroupID"] = clientGroupId;

                            dt.Rows[0]["Contacts"] = contacts;
                            dt.Rows[0]["ManagerID"] = managerId;

                            da.Update(dt);
                            _clientsDataTable.Clear();
                            ClientsDataAdapter.Fill(_clientsDataTable);
                            ClientsBindingSource.Position = ClientsBindingSource.Find("ClientID", clientId);
                        }
                    }
                }
            }
        }

        public void SaveShopAddresses()
        {
            ShopAddressesDataAdapter.Update(NewShopAddressesDataTable);
            ShopAddressesDataTable.Clear();
            ShopAddressesDataAdapter.Fill(ShopAddressesDataTable);
        }

        public void ExportShopAddressesFromExcel()
        {
            var table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("diler", typeof(string));
            table.Columns.Add("clientName", typeof(string));
            table.Columns.Add("Address", typeof(string));
            table.Columns.Add("City", typeof(string));
            table.Columns.Add("Phone", typeof(string));
            table.Columns.Add("Site", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("Country", typeof(string));
            table.TableName = "ImportedTable";

            var s = Clipboard.GetText();
            var lines = s.Split('\n');
            var data = new System.Collections.Generic.List<string>(lines);

            if (data.Count > 0 && string.IsNullOrWhiteSpace(data[data.Count - 1]))
            {
                data.RemoveAt(data.Count - 1);
            }

            foreach (var iterationRow in data)
            {
                var row = iterationRow;
                if (row.EndsWith("\r"))
                {
                    row = row.Substring(0, row.Length - "\r".Length);
                }

                var rowData = row.Split(new char[] { '\r', '\x09' });
                var newRow = table.NewRow();

                for (var i = 0; i < rowData.Length; i++)
                {
                    if (i >= table.Columns.Count) break;
                    if (rowData[i].Length > 0)
                        newRow[i] = rowData[i];
                }
                table.Rows.Add(newRow);
            }

            using (var da = new SqlDataAdapter(@"SELECT * FROM ShopAddresses",
                ConnectionStrings.ZOVReferenceConnectionString))
            {
                using (var cb = new SqlCommandBuilder(da))
                {
                    using (var dt = new DataTable())
                    {
                        da.Fill(dt);
                        for (var i = 0; i < table.Rows.Count; i++)
                        {
                            var row = dt.NewRow();
                            row["Country"] = table.Rows[i]["Country"];
                            row["City"] = table.Rows[i]["City"];
                            row["Address"] = table.Rows[i]["Address"];
                            row["Site"] = table.Rows[i]["Site"];
                            row["Phone"] = table.Rows[i]["Phone"];
                            row["Email"] = table.Rows[i]["Email"];
                            row["ClientID"] = table.Rows[i]["id"];

                            dt.Rows.Add(row);
                        }
                        da.Update(dt);
                    }
                }
            }
        }
    }

    public class ZovClientsToExcel
    {
        private int _pos01 = 0;

        private HSSFWorkbook _hssfworkbook;

        private HSSFFont _fConfirm;
        private HSSFFont _fHeader;
        private HSSFFont _fColumnName;
        private HSSFFont _fMainContent;
        private HSSFFont _fTotalInfo;

        private HSSFCellStyle _csConfirm;
        private HSSFCellStyle _csHeader;
        private HSSFCellStyle _csColumnName;
        private HSSFCellStyle _csMainContent;
        private HSSFCellStyle _csTotalInfo;

        public ZovClientsToExcel()
        {
            _hssfworkbook = new HSSFWorkbook();
            CreateFonts();
            CreateCellStyles();
        }

        private void CreateFonts()
        {
            _fConfirm = _hssfworkbook.CreateFont();
            _fConfirm.FontHeightInPoints = 12;
            _fConfirm.FontName = "Calibri";

            _fHeader = _hssfworkbook.CreateFont();
            _fHeader.FontHeightInPoints = 12;
            _fHeader.Boldweight = 12 * 256;
            _fHeader.FontName = "Calibri";

            _fColumnName = _hssfworkbook.CreateFont();
            _fColumnName.FontHeightInPoints = 12;
            _fColumnName.Boldweight = 12 * 256;
            _fColumnName.FontName = "Calibri";

            _fMainContent = _hssfworkbook.CreateFont();
            _fMainContent.FontHeightInPoints = 11;
            _fMainContent.FontName = "Calibri";

            _fTotalInfo = _hssfworkbook.CreateFont();
            _fTotalInfo.FontHeightInPoints = 11;
            _fTotalInfo.Boldweight = 11 * 256;
            _fTotalInfo.FontName = "Calibri";
        }

        private void CreateCellStyles()
        {
            _csConfirm = _hssfworkbook.CreateCellStyle();
            _csConfirm.SetFont(_fConfirm);

            _csHeader = _hssfworkbook.CreateCellStyle();
            _csHeader.SetFont(_fHeader);

            _csColumnName = _hssfworkbook.CreateCellStyle();
            _csColumnName.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _csColumnName.BottomBorderColor = HSSFColor.BLACK.index;
            _csColumnName.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _csColumnName.LeftBorderColor = HSSFColor.BLACK.index;
            _csColumnName.BorderRight = HSSFCellStyle.BORDER_THIN;
            _csColumnName.RightBorderColor = HSSFColor.BLACK.index;
            _csColumnName.BorderTop = HSSFCellStyle.BORDER_THIN;
            _csColumnName.TopBorderColor = HSSFColor.BLACK.index;
            _csColumnName.Alignment = HSSFCellStyle.ALIGN_CENTER;
            _csColumnName.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _csColumnName.WrapText = true;
            _csColumnName.SetFont(_fColumnName);

            _csMainContent = _hssfworkbook.CreateCellStyle();
            _csMainContent.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _csMainContent.BottomBorderColor = HSSFColor.BLACK.index;
            _csMainContent.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _csMainContent.LeftBorderColor = HSSFColor.BLACK.index;
            _csMainContent.BorderRight = HSSFCellStyle.BORDER_THIN;
            _csMainContent.RightBorderColor = HSSFColor.BLACK.index;
            _csMainContent.BorderTop = HSSFCellStyle.BORDER_THIN;
            _csMainContent.TopBorderColor = HSSFColor.BLACK.index;
            _csMainContent.VerticalAlignment = HSSFCellStyle.ALIGN_RIGHT;
            _csMainContent.SetFont(_fMainContent);

            _csTotalInfo = _hssfworkbook.CreateCellStyle();
            _csTotalInfo.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _csTotalInfo.BottomBorderColor = HSSFColor.BLACK.index;
            _csTotalInfo.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _csTotalInfo.LeftBorderColor = HSSFColor.BLACK.index;
            _csTotalInfo.BorderRight = HSSFCellStyle.BORDER_THIN;
            _csTotalInfo.RightBorderColor = HSSFColor.BLACK.index;
            _csTotalInfo.BorderTop = HSSFCellStyle.BORDER_THIN;
            _csTotalInfo.TopBorderColor = HSSFColor.BLACK.index;
            _csTotalInfo.Alignment = HSSFCellStyle.ALIGN_LEFT;
            _csTotalInfo.SetFont(_fTotalInfo);
        }

        public void ClearReport()
        {
            _hssfworkbook = new HSSFWorkbook();
            CreateFonts();
            CreateCellStyles();
        }

        public void Export(DataTable table1)
        {
            var sheet01 = _hssfworkbook.CreateSheet("Клиенты");
            sheet01.PrintSetup.PaperSize = (short)PaperSizeType.A4;
            sheet01.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet01.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet01.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet01.SetMargin(HSSFSheet.BottomMargin, (double).20);
            _pos01 = 0;

            var displayIndex = 0;
            {
                HSSFCell cell1 = null;
                for (var i = 0; i < table1.Columns.Count; i++)
                {
                    sheet01.SetColumnWidth(i, 15 * 256);
                    var colName = table1.Columns[i].ToString();

                    cell1 = sheet01.CreateRow(_pos01).CreateCell(displayIndex++);
                    cell1.SetCellValue(colName);
                    cell1.CellStyle = _csColumnName;
                    switch (colName)
                    {
                        case "ClientName":
                            sheet01.SetColumnWidth(i, 40 * 256);
                            cell1.SetCellValue("Клиент");
                            break;
                        case "ClientGroupName":
                            sheet01.SetColumnWidth(i, 40 * 256);
                            cell1.SetCellValue("Группа");
                            break;
                        case "Name":
                            sheet01.SetColumnWidth(i, 40 * 256);
                            cell1.SetCellValue("Менеджер");
                            break;
                    }
                }

                _pos01++;

                //Содержимое таблицы
                for (var x = 0; x < table1.Rows.Count; x++)
                {
                    for (var y = 0; y < table1.Columns.Count; y++)
                    {
                        var t = table1.Rows[x][y].GetType();

                        switch (t.Name)
                        {
                            case "Int32":
                                cell1 = sheet01.CreateRow(_pos01).CreateCell(y);
                                cell1.SetCellValue(Convert.ToInt32(table1.Rows[x][y]));
                                cell1.CellStyle = _csMainContent;
                                continue;
                            case "String":
                            case "DBNull":
                                cell1 = sheet01.CreateRow(_pos01).CreateCell(y);
                                cell1.SetCellValue(table1.Rows[x][y].ToString());
                                cell1.CellStyle = _csMainContent;
                                continue;
                        }
                    }

                    _pos01++;
                }
            }

            _pos01++;
            _pos01++;
        }

        public void SaveFile(string fileName, bool bOpenFile)
        {
            var tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            var file = new FileInfo(tempFolder + @"\" + fileName + ".xls");
            var j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + fileName + "(" + j++ + ").xls");
            }

            var newFile = new FileStream(file.FullName, FileMode.Create);
            _hssfworkbook.Write(newFile);
            newFile.Close();
            ClearReport();

            if (bOpenFile)
                System.Diagnostics.Process.Start(file.FullName);
        }
    }
}
