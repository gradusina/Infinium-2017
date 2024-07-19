using NPOI.HSSF.Record.Formula.Functions;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Management.Instrumentation;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Infinium.Modules.Marketing.Clients
{
    public class Clients
    {
        private PercentageDataGrid ClientsDataGrid = null;
        private PercentageDataGrid ContactsDataGrid = null;
        private PercentageDataGrid ShopAddressesDataGrid = null;

        public DataTable clientsExcluziveCountriesDt = null;
        public DataTable excluziveCountriesDt = null;

        public DataTable ClientRatesDataTable = null;
        public DataTable ClientsDt { get; private set; } = null;
        public DataTable NewCountriesDataTable = null;
        public DataTable CountriesDataTable = null;
        public DataTable ManagersDataTable = null;
        public DataTable GetManagersDataTable = null;
        public DataTable ClientGroupsDataTable = null;

        public bool NewClient = false;

        private DataTable FrontsPriceGroupsDataTable = null;
        private DataTable DecorPriceGroupsDataTable = null;
        public DataTable ContactsDataTable = null;
        public DataTable NewContactsDataTable = null;

        public DataTable ShopAddressesDataTable = null;
        public DataTable NewShopAddressesDataTable = null;

        public SqlDataAdapter CountriesDataAdapter = null;
        public SqlDataAdapter ClientRatesDataAdapter = null;
        public SqlDataAdapter ClientsDataAdapter = null;
        public SqlDataAdapter ShopAddressesDataAdapter = null;

        public SqlCommandBuilder CountriesCommandBuilder = null;
        public SqlCommandBuilder ClientRatesCommandBuilder = null;
        public SqlCommandBuilder ClientsCommandBuilder = null;
        public SqlCommandBuilder ShopAddressesCommandBuilder = null;

        public BindingSource excluziveCountriesBs = null;
        public BindingSource FrontsPriceGroupsBindingSource = null;
        public BindingSource DecorPriceGroupsBindingSource = null;
        public BindingSource ContactsBindingSource = null;
        public BindingSource NewContactsBindingSource = null;
        public BindingSource ClientRatesBindingSource = null;
        public BindingSource ClientsBindingSource = null;
        public BindingSource NewCountriesBindingSource = null;
        public BindingSource CountriesBindingSource = null;
        public BindingSource UsersBindingSource = null;
        public BindingSource ClientGroupsBindingSource = null;
        public BindingSource ShopAddressesBindingSource = null;
        public BindingSource NewShopAddressesBindingSource = null;

        public string ClientsBindingSourceDisplayMember = null;

        public string ClientsBindingSourceValueMember = null;


        public Clients(ref PercentageDataGrid tClientsDataGrid, ref PercentageDataGrid tContactsDataGrid, ref PercentageDataGrid tShopAddressesDataGrid)
        {
            ClientsDataGrid = tClientsDataGrid;
            ContactsDataGrid = tContactsDataGrid;
            ShopAddressesDataGrid = tShopAddressesDataGrid;

            Initialize();
        }

        private void Create()
        {
            prepareTranslit();
            clientsExcluziveCountriesDt = new DataTable();
            excluziveCountriesDt = new DataTable();

            FrontsPriceGroupsDataTable = new DataTable();
            DecorPriceGroupsDataTable = new DataTable();
            ClientRatesDataTable = new DataTable();
            ClientsDt = new DataTable();
            ClientGroupsDataTable = new DataTable();
            CountriesDataTable = new DataTable();
            ManagersDataTable = new DataTable();
            GetManagersDataTable = new DataTable();
            ShopAddressesDataTable = new DataTable();

            FrontsPriceGroupsBindingSource = new BindingSource();
            DecorPriceGroupsBindingSource = new BindingSource();
            ContactsBindingSource = new BindingSource();
            NewContactsBindingSource = new BindingSource();
            ClientRatesBindingSource = new BindingSource();
            ClientsBindingSource = new BindingSource();
            CountriesBindingSource = new BindingSource();
            UsersBindingSource = new BindingSource();
            ClientGroupsBindingSource = new BindingSource();
            ShopAddressesBindingSource = new BindingSource();
            NewShopAddressesBindingSource = new BindingSource();
        }

        private void Fill()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM ClientsExcluziveCountries order by clientId",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(clientsExcluziveCountriesDt);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Countries ORDER BY Name",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(excluziveCountriesDt);
            }
            var column1 = new DataColumn("check")
            {
                DataType = typeof(bool),
                DefaultValue = 0
            };
            excluziveCountriesDt.Columns.Add(column1);

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsPriceGroups",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(FrontsPriceGroupsDataTable);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorPriceGroups",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(DecorPriceGroupsDataTable);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ManagerID, Name, ShortName FROM ClientsManagers ORDER BY Name", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ManagersDataTable);
                {
                    DataRow NewRow = ManagersDataTable.NewRow();
                    NewRow["ManagerID"] = 0;
                    NewRow["Name"] = "-";
                    NewRow["ShortName"] = "-";
                    ManagersDataTable.Rows.InsertAt(NewRow, 0);
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM ClientsManagers ORDER BY Name", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(GetManagersDataTable);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Countries ORDER BY Name",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CountriesDataTable);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM ClientGroups ORDER BY ClientGroupName",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientGroupsDataTable);
            }

            ClientRatesDataAdapter = new SqlDataAdapter("SELECT TOP 0 * FROM ClientRates ORDER BY Date",
                ConnectionStrings.MarketingReferenceConnectionString);
            ClientRatesCommandBuilder = new SqlCommandBuilder(ClientRatesDataAdapter);
            ClientRatesDataAdapter.Fill(ClientRatesDataTable);

            ClientsDataAdapter = new SqlDataAdapter("SELECT * FROM Clients ORDER BY ClientName",
                ConnectionStrings.MarketingReferenceConnectionString);
            ClientsCommandBuilder = new SqlCommandBuilder(ClientsDataAdapter);
            ClientsDataAdapter.Fill(ClientsDt);

            var column2 = new DataColumn("ExcluziveCountries")
            {
                DataType = typeof(string),
                DefaultValue = string.Empty
            };
            ClientsDt.Columns.Add(column2);
            ClientsDt.Columns.Add(new DataColumn("USD", typeof(decimal)));
            ClientsDt.Columns.Add(new DataColumn("RUB", typeof(decimal)));
            ClientsDt.Columns.Add(new DataColumn("BYN", typeof(decimal)));

            FillСlientsExcluziveCountries();
            FillClientRates();

            ShopAddressesDataAdapter = new SqlDataAdapter("SELECT TOP 0 * FROM ShopAddresses ORDER BY Address",
                ConnectionStrings.MarketingReferenceConnectionString);
            ShopAddressesCommandBuilder = new SqlCommandBuilder(ShopAddressesDataAdapter);
            ShopAddressesDataAdapter.Fill(ShopAddressesDataTable);
        }

        private void Binding()
        {
            excluziveCountriesBs = new BindingSource()
            {
                DataSource = excluziveCountriesDt
            };

            FrontsPriceGroupsBindingSource.DataSource = FrontsPriceGroupsDataTable;
            DecorPriceGroupsBindingSource.DataSource = DecorPriceGroupsDataTable;

            ClientRatesBindingSource.DataSource = ClientRatesDataTable;
            ClientsBindingSource.DataSource = ClientsDt;
            CountriesBindingSource.DataSource = CountriesDataTable;
            UsersBindingSource.DataSource = ManagersDataTable;
            ClientGroupsBindingSource.DataSource = ClientGroupsDataTable;
            ContactsBindingSource.DataSource = ContactsDataTable;
            ShopAddressesBindingSource.DataSource = ShopAddressesDataTable;

            ClientsDataGrid.DataSource = ClientsBindingSource;
            ContactsDataGrid.DataSource = ContactsBindingSource;
            ShopAddressesDataGrid.DataSource = ShopAddressesBindingSource;


            ClientsBindingSourceDisplayMember = "ClientName";

            ClientsBindingSourceValueMember = "ClientID";
        }

        private void SetClientsDataGrid()
        {
            ClientsDataGrid.Columns.Add(CountryColumn);
            ClientsDataGrid.Columns.Add(ManagerColumn);
            ClientsDataGrid.Columns.Add(ClientGroupColumn);

            foreach (DataGridViewColumn Column in ClientsDataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            ClientsDataGrid.Columns["CountryID"].Visible = false;
            ClientsDataGrid.Columns["ManagerID"].Visible = false;
            ClientsDataGrid.Columns["Login"].Visible = false;
            ClientsDataGrid.Columns["Password"].Visible = false;
            ClientsDataGrid.Columns["Contacts"].Visible = false;
            ClientsDataGrid.Columns["ClientGroupID"].Visible = false;
            ClientsDataGrid.Columns["Online"].Visible = false;
            ClientsDataGrid.Columns["OnlineRefreshDateTime"].Visible = false;
            ClientsDataGrid.Columns["IdleTime"].Visible = false;
            ClientsDataGrid.Columns["TopMost"].Visible = false;
            ClientsDataGrid.Columns["TopModule"].Visible = false;
            ClientsDataGrid.Columns["Accountant"].Visible = false;
            ClientsDataGrid.Columns["Director"].Visible = false;
            ClientsDataGrid.Columns["DischargeAddress"].Visible = false;
            ClientsDataGrid.Columns["ClientFolderName"].Visible = false;

            ClientsDataGrid.Columns["Enabled"].HeaderText = "Активен";
            ClientsDataGrid.Columns["PriceGroup"].HeaderText = "Ценовая\r\nгруппа";
            ClientsDataGrid.Columns["DelayOfPayment"].HeaderText = "Отсрочка";
            ClientsDataGrid.Columns["ClientID"].HeaderText = "ID";
            ClientsDataGrid.Columns["ClientName"].HeaderText = "Название организации";
            ClientsDataGrid.Columns["Login"].HeaderText = "Логин";
            //ClientsDataGrid.Columns["Country"].HeaderText = "Страна";
            ClientsDataGrid.Columns["City"].HeaderText = "Город";
            ClientsDataGrid.Columns["Email"].HeaderText = "E-mail";
            ClientsDataGrid.Columns["Site"].HeaderText = "Сайт";
            ClientsDataGrid.Columns["UNN"].HeaderText = "УНН";
            ClientsDataGrid.Columns["ExcluziveCountries"].HeaderText = "Страны по\r\nэксклюзиву";
            ClientsDataGrid.Columns["NonStandard"].HeaderText = "Учет\r\nнестандарта";

            ClientsDataGrid.AutoGenerateColumns = false;

            ClientsDataGrid.Columns["PriceGroup"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["DelayOfPayment"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["Enabled"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["ClientID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ClientsDataGrid.Columns["ClientID"].Width = 50;
            ClientsDataGrid.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["ClientName"].MinimumWidth = 150;
            ClientsDataGrid.Columns["ManagerColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["ManagerColumn"].MinimumWidth = 150;
            ClientsDataGrid.Columns["ClientGroupColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["ClientGroupColumn"].MinimumWidth = 150;
            ClientsDataGrid.Columns["CountryColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["CountryColumn"].MinimumWidth = 100;
            ClientsDataGrid.Columns["City"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["City"].MinimumWidth = 100;
            ClientsDataGrid.Columns["Email"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["Email"].MinimumWidth = 100;
            ClientsDataGrid.Columns["UNN"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["UNN"].MinimumWidth = 100;
            ClientsDataGrid.Columns["Site"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["Site"].MinimumWidth = 70;
            ClientsDataGrid.Columns["ExcluziveCountries"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ClientsDataGrid.Columns["ExcluziveCountries"].MinimumWidth = 70;
            //ClientsDataGrid.Columns["Login"].MinimumWidth = 60;
            ClientsDataGrid.Columns["NonStandard"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ClientsDataGrid.Columns["NonStandard"].Width = 120;
            ClientsDataGrid.Columns["USD"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ClientsDataGrid.Columns["USD"].Width = 50;
            ClientsDataGrid.Columns["RUB"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ClientsDataGrid.Columns["RUB"].Width = 50;
            ClientsDataGrid.Columns["BYN"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ClientsDataGrid.Columns["BYN"].Width = 50;

            int DisplayIndex = 0;
            ClientsDataGrid.Columns["ClientID"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["ClientName"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["ClientGroupColumn"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["ManagerColumn"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["Login"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["CountryColumn"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["City"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["UNN"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["NonStandard"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["DelayOfPayment"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["PriceGroup"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["Enabled"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["USD"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["RUB"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["BYN"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["Site"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["Email"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["ExcluziveCountries"].DisplayIndex = DisplayIndex++;

            ClientsDataGrid.Columns["ManagerColumn"].ReadOnly = true;
            ClientsDataGrid.Columns["ClientGroupColumn"].ReadOnly = true;
            ClientsDataGrid.Columns["ClientID"].ReadOnly = true;
            ClientsDataGrid.Columns["ClientName"].ReadOnly = true;
            ClientsDataGrid.Columns["Login"].ReadOnly = true;
            ClientsDataGrid.Columns["CountryColumn"].ReadOnly = true;
            ClientsDataGrid.Columns["City"].ReadOnly = true;
            ClientsDataGrid.Columns["Email"].ReadOnly = true;
            ClientsDataGrid.Columns["UNN"].ReadOnly = true;
            ClientsDataGrid.Columns["Site"].ReadOnly = true;
            ClientsDataGrid.Columns["NonStandard"].ReadOnly = true;
            ClientsDataGrid.Columns["DelayOfPayment"].ReadOnly = true;
            //ClientsDataGrid.Columns["ManagerColumn"].ReadOnly = false;
        }

        private DataGridViewComboBoxColumn ManagerColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "ManagerColumn",
                    HeaderText = "Кто ведёт",
                    DataPropertyName = "ManagerID",

                    DataSource = new DataView(ManagersDataTable),
                    ValueMember = "ManagerID",
                    DisplayMember = "Name",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return Column;
            }
        }

        private DataGridViewComboBoxColumn ClientGroupColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "ClientGroupColumn",
                    HeaderText = "Группа",
                    DataPropertyName = "ClientGroupID",

                    DataSource = new DataView(ClientGroupsDataTable),
                    ValueMember = "ClientGroupID",
                    DisplayMember = "ClientGroupName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return Column;
            }
        }

        private DataGridViewComboBoxColumn CountryColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "CountryColumn",
                    HeaderText = "Страна",
                    DataPropertyName = "CountryID",

                    DataSource = new DataView(CountriesDataTable),
                    ValueMember = "CountryID",
                    DisplayMember = "Name",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return Column;
            }
        }

        private void SetContactsDataGrid()
        {
            ContactsDataGrid.Columns["Name"].HeaderText = "Имя";
            ContactsDataGrid.Columns["Position"].HeaderText = "Должность";
            ContactsDataGrid.Columns["Phone"].HeaderText = "Телефон";
            ContactsDataGrid.Columns["Email"].HeaderText = "E-mail";
            ContactsDataGrid.Columns["ICQ"].HeaderText = "ICQ";
            ContactsDataGrid.Columns["Notes"].HeaderText = "Примечание";

            ContactsDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            ShopAddressesDataGrid.Columns["ShopAddressID"].Visible = false;
            ShopAddressesDataGrid.Columns["ClientID"].Visible = false;
            ShopAddressesDataGrid.Columns["Address"].HeaderText = "Адрес";
            ShopAddressesDataGrid.Columns["City"].HeaderText = "Город";
            ShopAddressesDataGrid.Columns["Country"].HeaderText = "Страна";
            ShopAddressesDataGrid.Columns["Lat"].HeaderText = "Широта";
            ShopAddressesDataGrid.Columns["Long"].HeaderText = "Долгота";
            ShopAddressesDataGrid.Columns["Worktime"].HeaderText = "Время работы";
            ShopAddressesDataGrid.Columns["Name"].HeaderText = "Название";
            ShopAddressesDataGrid.Columns["Site"].HeaderText = "Сайт";
            ShopAddressesDataGrid.Columns["IsFronts"].HeaderText = "Фасады";
            ShopAddressesDataGrid.Columns["IsProfile"].HeaderText = "Погонаж";
            ShopAddressesDataGrid.Columns["IsFurniture"].HeaderText = "Мебель";
            ShopAddressesDataGrid.Columns["Email"].HeaderText = "Email";
            ShopAddressesDataGrid.Columns["Phone"].HeaderText = "Телефон";
            ShopAddressesDataGrid.Columns["WebSite"].HeaderText = "Сайт";

            ShopAddressesDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            ShopAddressesDataGrid.Columns["Address"].DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;
            ShopAddressesDataGrid.Columns["Address"].DefaultCellStyle.Font = new System.Drawing.Font("SEGOE UI", 13.0F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Pixel, ((byte)(0)));
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

        private string GetContactsXML(DataTable DT)
        {
            StringWriter SW = new StringWriter();
            DT.WriteXml(SW);

            return SW.ToString();
        }

        public void FillContactsDataTable(int ClientID)
        {
            ContactsDataTable.Clear();

            DataRow[] Rows = ClientsDt.Select("ClientID = " + ClientID);

            if (Rows[0]["Contacts"].ToString().Length == 0)
                return;

            string ContactXML = Rows[0]["Contacts"].ToString();

            using (StringReader SR = new StringReader(ContactXML))
            {
                ContactsDataTable.ReadXml(SR);
            }
        }

        public void FillShopAddressesDataTable(int ClientID)
        {
            ShopAddressesDataAdapter.Dispose();
            ShopAddressesCommandBuilder.Dispose();
            ShopAddressesDataAdapter = new SqlDataAdapter("SELECT * FROM ShopAddresses WHERE ClientID=" + ClientID + " ORDER BY Address",
                ConnectionStrings.MarketingReferenceConnectionString);
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

        private bool GetManagerInfo(int ManagerID, ref string Name, ref string Phone, ref string Email, ref string Skype)
        {
            DataRow[] Rows = GetManagersDataTable.Select("ManagerID = " + ManagerID);
            if (Rows.Count() > 0)
            {
                Name = Rows[0]["Name"].ToString();
                Phone = Rows[0]["Phone"].ToString();
                Email = Rows[0]["Email"].ToString();
                Skype = Rows[0]["Skype"].ToString();
                if (Phone.Count() == 0 && Email.Count() == 0 && Skype.Count() == 0)
                    return false;
                return true;
            }
            else
                return false;
        }


        public string GetClientName(int ClientID)
        {
            DataRow[] Row = ClientsDt.Select("ClientID = " + ClientID);

            return Row[0]["ClientName"].ToString();
        }

        public string GetCountryName(int CountryID)
        {
            DataRow[] Row = CountriesDataTable.Select("CountryID = " + CountryID);

            return Row[0]["Name"].ToString();
        }

        public void CreateNewContactsDataTable(ref PercentageDataGrid DataGrid, ref PercentageDataGrid tShopAddressesDataGrid)
        {
            if (NewContactsDataTable == null)
            {
                NewContactsDataTable = new DataTable();
                NewContactsDataTable = ContactsDataTable.Clone();
            }
            else
                NewContactsDataTable.Clear();

            NewContactsBindingSource.DataSource = NewContactsDataTable;
            DataGrid.DataSource = NewContactsBindingSource;

            DataGrid.Columns["Name"].HeaderText = "Имя";
            DataGrid.Columns["Position"].HeaderText = "Должность";
            DataGrid.Columns["Phone"].HeaderText = "Телефон";
            DataGrid.Columns["Email"].HeaderText = "E-mail";
            DataGrid.Columns["ICQ"].HeaderText = "ICQ";
            DataGrid.Columns["Notes"].HeaderText = "Примечание";

            DataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            if (NewShopAddressesDataTable == null)
            {
                NewShopAddressesDataTable = new DataTable();
                NewShopAddressesDataTable = ShopAddressesDataTable.Clone();
            }
            else
                NewShopAddressesDataTable.Clear();

            NewShopAddressesBindingSource.DataSource = NewShopAddressesDataTable;
            tShopAddressesDataGrid.DataSource = NewShopAddressesBindingSource;

            tShopAddressesDataGrid.Columns["ShopAddressID"].Visible = false;
            tShopAddressesDataGrid.Columns["ClientID"].Visible = false;
            tShopAddressesDataGrid.Columns["Address"].HeaderText = "Адрес";
            tShopAddressesDataGrid.Columns["City"].HeaderText = "Город";
            tShopAddressesDataGrid.Columns["Country"].HeaderText = "Страна";
            tShopAddressesDataGrid.Columns["Lat"].HeaderText = "Широта";
            tShopAddressesDataGrid.Columns["Long"].HeaderText = "Долгота";
            tShopAddressesDataGrid.Columns["Worktime"].HeaderText = "Время работы";
            tShopAddressesDataGrid.Columns["Name"].HeaderText = "Название";
            tShopAddressesDataGrid.Columns["Site"].HeaderText = "Сайт";
            tShopAddressesDataGrid.Columns["IsFronts"].HeaderText = "Фасады";
            tShopAddressesDataGrid.Columns["IsProfile"].HeaderText = "Погонаж";
            tShopAddressesDataGrid.Columns["IsFurniture"].HeaderText = "Мебель";
            tShopAddressesDataGrid.Columns["Email"].HeaderText = "Email";
            tShopAddressesDataGrid.Columns["Phone"].HeaderText = "Телефон";
            tShopAddressesDataGrid.Columns["WebSite"].HeaderText = "Сайт";

            tShopAddressesDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            //tShopAddressesDataGrid.Columns["Address"].DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;
            //tShopAddressesDataGrid.Columns["Address"].DefaultCellStyle.Font = new System.Drawing.Font("SEGOE UI", 13.0F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Pixel, ((byte)(0)));
        }


        public void AddClient(string Name, int CountryID, string City, int ClientGroupID, string Site, string Email, int ManagerID, string UNN,
                              int NonStandard, decimal PriceGroup, int DelayOfPayment, bool Enabled)
        {
            string Contacts = GetContactsXML(NewContactsDataTable);

            string Login = GetTranslit(Name);
            DataRow Row = ClientsDt.NewRow();
            Row["ManagerID"] = ManagerID;
            Row["Login"] = Login;
            Row["Password"] = "b59c67bf196a4758191e42f76670ceba";
            Row["ClientName"] = Name;
            Row["ClientFolderName"] = Name.Replace('"', ' ');
            Row["CountryID"] = CountryID;
            Row["ClientGroupID"] = ClientGroupID;
            Row["City"] = City;
            Row["Token"] = GenAccessToken(Name);

            if (Email.Length > 0)
                Row["Email"] = Email;

            if (Contacts.Length > 0)
                Row["Site"] = Site;

            if (UNN != string.Empty)
                Row["UNN"] = UNN;
            else
                Row["UNN"] = DBNull.Value;

            Row["Contacts"] = Contacts;
            Row["NonStandard"] = NonStandard;
            Row["PriceGroup"] = PriceGroup;
            Row["DelayOfPayment"] = DelayOfPayment;
            Row["Enabled"] = Enabled;

            ClientsDt.Rows.Add(Row);
            ClientsDataAdapter.Update(ClientsDt);
            ClientsDt.Clear();
            ClientsDataAdapter.Fill(ClientsDt);
        }

        public void EditClient(ref string ClientsName, ref int ClientsCountryID, ref string ClientsCity, ref int ClientGroupID,
                               ref string ClientsEmail, ref string ClientsSite, ref string UNN, ref int ManagerID, ref decimal PriceGroup, ref int DelayOfPayment, ref bool Enabled)
        {
            ClientsName = ((DataRowView)ClientsBindingSource.Current)["ClientName"].ToString();
            ClientsCountryID = Convert.ToInt32(((DataRowView)ClientsBindingSource.Current)["CountryID"]);
            ClientGroupID = Convert.ToInt32(((DataRowView)ClientsBindingSource.Current)["ClientGroupID"]);
            ClientsCity = ((DataRowView)ClientsBindingSource.Current)["City"].ToString();
            ClientsEmail = ((DataRowView)ClientsBindingSource.Current)["Email"].ToString();
            ClientsSite = ((DataRowView)ClientsBindingSource.Current)["Site"].ToString();
            UNN = ((DataRowView)ClientsBindingSource.Current)["UNN"].ToString();
            ManagerID = Convert.ToInt32(((DataRowView)ClientsBindingSource.Current)["ManagerID"]);
            PriceGroup = Convert.ToDecimal(((DataRowView)ClientsBindingSource.Current)["PriceGroup"]);
            DelayOfPayment = Convert.ToInt32(((DataRowView)ClientsBindingSource.Current)["DelayOfPayment"]);
            Enabled = Convert.ToBoolean(((DataRowView)ClientsBindingSource.Current)["Enabled"]);

            NewContactsDataTable.Clear();
            NewContactsDataTable = ContactsDataTable.Copy();

            NewContactsBindingSource.DataSource = NewContactsDataTable;

            NewShopAddressesDataTable.Clear();
            NewShopAddressesDataTable = ShopAddressesDataTable.Copy();
            NewShopAddressesBindingSource.DataSource = NewShopAddressesDataTable;
        }

        private void AddСlientsExcluziveCountries(int clientId)
        {
            for (var i = 0; i < excluziveCountriesDt.Rows.Count; i++)
            {
                var check = Convert.ToBoolean(excluziveCountriesDt.Rows[i]["check"]);

                if (!check)
                    continue;
                
                var countryId = Convert.ToInt32(excluziveCountriesDt.Rows[i]["countryId"]);

                var rows = clientsExcluziveCountriesDt.Select($"clientId={clientId} and countryId={countryId}");
                if (rows.Any()) continue;

                var row = clientsExcluziveCountriesDt.NewRow();
                row["clientId"] = clientId;
                row["countryId"] = countryId;
                clientsExcluziveCountriesDt.Rows.Add(row);
            }
        }

        private void RemoveСlientsExcluziveCountries(int clientId)
        {
            for (var i = excluziveCountriesDt.Rows.Count - 1; i >= 0; i--)
            {
                var check = Convert.ToBoolean(excluziveCountriesDt.Rows[i]["check"]);

                if (check)
                    continue;
                
                var countryId = Convert.ToInt32(excluziveCountriesDt.Rows[i]["countryId"]);

                var rows = clientsExcluziveCountriesDt.Select($"clientId={clientId} and countryId={countryId}");
                if (!rows.Any()) continue;

                rows[0].Delete();
            }
        }

        private void FillСlientsExcluziveCountries()
        {
            for (var i = ClientsDt.Rows.Count - 1; i >= 0; i--)
            {

                var rows = clientsExcluziveCountriesDt.Select($"clientId={Convert.ToInt32(ClientsDt.Rows[i]["clientId"])}");
                if (!rows.Any()) continue;

                var countryId = Convert.ToInt32(rows[0]["countryId"]);
                var sb = new StringBuilder(GetCountryName(Convert.ToInt32(rows[0]["countryId"])) + ", ");

                for (var j = 1; j < rows.Length; j++)
                {
                    if (Convert.ToInt32(rows[j]["countryId"]) == countryId)
                        continue;
                    sb.Append(GetCountryName(Convert.ToInt32(rows[j]["countryId"])) + ", ");
                }

                var countriesList = sb.ToString();
                if (countriesList.Length > 0)
                    countriesList = countriesList.Substring(0, countriesList.Length - 2);
                ClientsDt.Rows[i]["ExcluziveCountries"] = countriesList;
            }
        }

        public void GetСlientsExcluziveCountries(int clientId)
        {
            for (var i = 0; i < excluziveCountriesDt.Rows.Count; i++)
            {
                excluziveCountriesDt.Rows[i]["check"] = false;
            }
            for (var i = 0; i < clientsExcluziveCountriesDt.Rows.Count; i++)
            {
                if (clientId != Convert.ToInt32(clientsExcluziveCountriesDt.Rows[i]["clientId"]))
                    continue;
                
                var countryId = Convert.ToInt32(clientsExcluziveCountriesDt.Rows[i]["countryId"]);

                var rows = excluziveCountriesDt.Select($"countryId={countryId}");
                if (!rows.Any()) continue;

                rows[0]["check"] = true;
            }
        }

        public void SaveСlientsExcluziveCountries(int clientId)
        {
            AddСlientsExcluziveCountries(clientId);
            RemoveСlientsExcluziveCountries(clientId);

            const string selectCommand = @"SELECT * FROM ClientsExcluziveCountries order by clientId";
            using (var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    da.Update(clientsExcluziveCountriesDt);
                    clientsExcluziveCountriesDt.Clear();
                    da.Fill(clientsExcluziveCountriesDt);
                }
            }
        }

        public void SaveClient(string Name, int CountryID, string City, int ClientGroupID, string Site, string Email, int ManagerID, string UNN,
                               int NonStandard, decimal PriceGroup, int DelayOfPayment, bool Enabled, int ClientID)
        {
            string Contacts = GetContactsXML(NewContactsDataTable);

            int index = ClientsBindingSource.Position;

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM Clients WHERE ClientID = " + ClientID + " ORDER BY ClientName",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);
                        if (DT.Rows.Count > 0)
                        {
                            if (DT.Rows[0]["Token"] == DBNull.Value)
                                DT.Rows[0]["Token"] = GenAccessToken(Name);

                            DT.Rows[0]["ClientName"] = Name;
                            DT.Rows[0]["CountryID"] = CountryID;
                            DT.Rows[0]["ClientGroupID"] = ClientGroupID;
                            DT.Rows[0]["City"] = City;

                            if (Email.Length > 0)
                                DT.Rows[0]["Email"] = Email;

                            if (Contacts.Length > 0)
                                DT.Rows[0]["Site"] = Site;
                            if (UNN != string.Empty)
                                DT.Rows[0]["UNN"] = UNN;
                            else
                                DT.Rows[0]["UNN"] = DBNull.Value;

                            DT.Rows[0]["Contacts"] = Contacts;
                            DT.Rows[0]["NonStandard"] = NonStandard;
                            DT.Rows[0]["PriceGroup"] = PriceGroup;
                            DT.Rows[0]["DelayOfPayment"] = DelayOfPayment;
                            DT.Rows[0]["Enabled"] = Enabled;

                            //int OldManagerID = 0;
                            //if (DT.Rows[0]["ManagerID"] != DBNull.Value)
                            //{
                            //    OldManagerID = Convert.ToInt32(DT.Rows[0]["ManagerID"]);
                            //    if (OldManagerID != ManagerID && ManagerID != 0)
                            //    {
                            //    }
                            //}
                            DT.Rows[0]["ManagerID"] = ManagerID;
                            DA.Update(DT);
                            ClientsDt.Clear();
                            ClientsDataAdapter.Fill(ClientsDt);
                            FillСlientsExcluziveCountries();
                            ClientsBindingSource.Position = ClientsBindingSource.Find("ClientID", ClientID);
                        }
                    }
                }
            }
        }

        private string SendNotifyToEmail(string ClientEmail, string Name, string Phone, string Email, string Skype)
        {
            string result = "Уведомление на почту клиента отправлено";

            //string AccountPassword = "7026Gradus0462";
            //var AccountPassword = "foqwsulbjiuslnue";
            var AccountPassword = "lfbeecgxvmwvzlna";
            string SenderEmail = "infiniumdevelopers@gmail.com";


            string to = ClientEmail;
            string from = SenderEmail;

            if (to.Length == 0)
            {
                result = "У клиента не указан Email. Отправка отчета невозможна";
                MessageBox.Show(result);
                return result;
            }

            string messageBody = "Уважаемые клиенты! Сообщаем Вам о назначении нового менеджера, курирующего Вашу компанию - " + Name + ".\r\n";
            if (Phone.Count() > 0 || Email.Count() > 0 || Skype.Count() > 0)
                messageBody += "Обращайтесь по следующим контактам:\r\n";
            if (Phone.Count() > 0)
                messageBody += "Тел: " + Phone + "\r\n";
            if (Email.Count() > 0)
                messageBody += "Эл. почта: " + Email + "\r\n";
            if (Skype.Count() > 0)
                messageBody += "Скайп: " + Skype;

            to = to.Replace(';', ',');
            using (MailMessage message = new MailMessage(from, to))
            {
                message.Subject = "О смене ответственного лица";
                message.Body = messageBody;
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
                {
                    EnableSsl = true,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(SenderEmail, AccountPassword)
                };
                try
                {
                    client.Send(message);
                }

                catch (ArgumentException ex)
                {
                    MessageBox.Show("ArgumentException\r\n" + ex.Message);
                    result = "ArgumentException\r\n" + ex.Message;
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show("InvalidOperationException\r\n" + ex.Message);
                    result = "InvalidOperationException\r\n" + ex.Message;
                }
                catch (SmtpException ex)
                {
                    MessageBox.Show("SmtpException\r\n" + ex.Message);
                    result = "SmtpException\r\n" + ex.Message;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception\r\n" + ex.Message);
                    result = "Exception\r\n" + ex.Message;
                }

                client.Dispose();
            }
            return result;
        }

        public string NotifyClient(string ClientEmail, int ManagerID)
        {
            string ManagerName = string.Empty;
            string ManagerPhone = string.Empty;
            string ManagerEmail = string.Empty;
            string ManagerSkype = string.Empty;
            string result = string.Empty;
            if (!GetManagerInfo(ManagerID, ref ManagerName, ref ManagerPhone, ref ManagerEmail, ref ManagerSkype))
                result = "Нет сведений о менеджере. Заполните информацию в модуле \r\n" +
                        "\"Редактор менеджеров\" и повторите отправку уведомления";
            else
                result = SendNotifyToEmail(ClientEmail, ManagerName, ManagerPhone, ManagerEmail, ManagerSkype);

            return result;
        }

        private string GetMD5(string text)
        {
            using (System.Security.Cryptography.MD5 Hasher = System.Security.Cryptography.MD5.Create())
            {
                byte[] data = Hasher.ComputeHash(Encoding.Default.GetBytes(text));

                StringBuilder sBuilder = new StringBuilder();

                //преобразование в HEX
                for (int i = 0; i < data.Length; i++)
                {
                    sBuilder.Append(data[i].ToString("x2"));
                }

                return sBuilder.ToString();
            }
        }

        public string GenAccessToken(string Name)
        {
            string AccessToken = string.Empty;
            if (Name.Length > 0)
            {
                Guid g = Guid.NewGuid();
                string GuidString = Convert.ToBase64String(g.ToByteArray());
                AccessToken = GetMD5(Name) + GetMD5(GuidString);
            }
            return AccessToken;
        }

        public void SaveShopAddresses()
        {
            ShopAddressesDataAdapter.Update(NewShopAddressesDataTable);
            ShopAddressesDataTable.Clear();
            ShopAddressesDataAdapter.Fill(ShopAddressesDataTable);
        }

        public void GetClientRates(int ClientID)
        {
            ClientRatesDataTable.Clear();
            ClientRatesDataAdapter = new SqlDataAdapter("SELECT * FROM ClientRates WHERE ClientID=" + ClientID + " ORDER BY Date",
                ConnectionStrings.MarketingReferenceConnectionString);
            ClientRatesCommandBuilder = new SqlCommandBuilder(ClientRatesDataAdapter);
            ClientRatesDataTable.Clear();
            ClientRatesDataAdapter.Fill(ClientRatesDataTable);
        }

        public void SaveClientRates()
        {
            CheckClientRateTime();
            ClientRatesDataAdapter.Update(ClientRatesDataTable);
            ClientRatesDataTable.Clear();
            ClientRatesDataAdapter.Fill(ClientRatesDataTable);
        }

        private void CheckClientRateTime()
        {
            for (int i = 0; i < ClientRatesDataTable.Rows.Count; i++)
            {
                if (ClientRatesDataTable.Rows[i].RowState == DataRowState.Deleted)
                    continue;
                if (ClientRatesDataTable.Rows[i]["Date"] != DBNull.Value)
                {
                    DateTime dateTime = Convert.ToDateTime(ClientRatesDataTable.Rows[i]["Date"]);
                    string timeString = dateTime.ToShortTimeString();
                    if (timeString == "0:00")
                        continue;
                    else
                    {
                        ClientRatesDataTable.Rows[i]["Date"] = dateTime.ToShortDateString() + " 00:00:00";
                    }
                }
            }
        }

        public void GetCountries()
        {
            if (NewCountriesDataTable == null)
            {
                NewCountriesDataTable = new DataTable();
                NewCountriesBindingSource = new BindingSource
                {
                    DataSource = NewCountriesDataTable
                };
            }
            NewCountriesDataTable.Clear();
            CountriesDataAdapter = new SqlDataAdapter("SELECT * FROM Countries ORDER BY Name",
                ConnectionStrings.CatalogConnectionString);
            CountriesCommandBuilder = new SqlCommandBuilder(CountriesDataAdapter);
            CountriesDataAdapter.Fill(NewCountriesDataTable);
        }

        public void UpdateCountries()
        {
            CountriesDataTable.Clear();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Countries ORDER BY Name",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CountriesDataTable);
            }
        }

        public void SaveCountries()
        {
            CountriesDataAdapter.Update(NewCountriesDataTable);
            NewCountriesDataTable.Clear();
            CountriesDataAdapter.Fill(NewCountriesDataTable);
        }

        public void AddCountry(string Name)
        {
            DataRow Row = NewCountriesDataTable.NewRow();
            Row["Name"] = Name;
            NewCountriesDataTable.Rows.Add(Row);
        }

        private Dictionary<string, string> transliter = new Dictionary<string, string>();

        private void prepareTranslit()
        {
            transliter.Add("а", "a");
            transliter.Add("б", "b");
            transliter.Add("в", "v");
            transliter.Add("г", "g");
            transliter.Add("д", "d");
            transliter.Add("е", "e");
            transliter.Add("ё", "yo");
            transliter.Add("ж", "zh");
            transliter.Add("з", "z");
            transliter.Add("и", "i");
            transliter.Add("й", "j");
            transliter.Add("к", "k");
            transliter.Add("л", "l");
            transliter.Add("м", "m");
            transliter.Add("н", "n");
            transliter.Add("о", "o");
            transliter.Add("п", "p");
            transliter.Add("р", "r");
            transliter.Add("с", "s");
            transliter.Add("т", "t");
            transliter.Add("у", "u");
            transliter.Add("ф", "f");
            transliter.Add("х", "h");
            transliter.Add("ц", "c");
            transliter.Add("ч", "ch");
            transliter.Add("ш", "sh");
            transliter.Add("щ", "sch");
            transliter.Add("ъ", "j");
            transliter.Add("ы", "i");
            transliter.Add("ь", "j");
            transliter.Add("э", "e");
            transliter.Add("ю", "yu");
            transliter.Add("я", "ya");
            transliter.Add("А", "A");
            transliter.Add("Б", "B");
            transliter.Add("В", "V");
            transliter.Add("Г", "G");
            transliter.Add("Д", "D");
            transliter.Add("Е", "E");
            transliter.Add("Ё", "Yo");
            transliter.Add("Ж", "Zh");
            transliter.Add("З", "Z");
            transliter.Add("И", "I");
            transliter.Add("Й", "J");
            transliter.Add("К", "K");
            transliter.Add("Л", "L");
            transliter.Add("М", "M");
            transliter.Add("Н", "N");
            transliter.Add("О", "O");
            transliter.Add("П", "P");
            transliter.Add("Р", "R");
            transliter.Add("С", "S");
            transliter.Add("Т", "T");
            transliter.Add("У", "U");
            transliter.Add("Ф", "F");
            transliter.Add("Х", "H");
            transliter.Add("Ц", "C");
            transliter.Add("Ч", "Ch");
            transliter.Add("Ш", "Sh");
            transliter.Add("Щ", "Sch");
            transliter.Add("Ъ", "J");
            transliter.Add("Ы", "I");
            transliter.Add("Ь", "J");
            transliter.Add("Э", "E");
            transliter.Add("Ю", "Yu");
            transliter.Add("Я", "Ya");
        }

        public string GetTranslit(string sourceText)
        {
            StringBuilder ans = new StringBuilder();
            for (int i = 0; i < sourceText.Length; i++)
            {
                if (transliter.ContainsKey(sourceText[i].ToString()))
                {
                    ans.Append(transliter[sourceText[i].ToString()]);
                }
                else
                {
                    ans.Append(sourceText[i].ToString());
                }
            }
            return ans.ToString();
        }

        public void Send(int ClientID, string ClientEmail, string Login)
        {
            //string AccountPassword = "7026Gradus0462";
            //var AccountPassword = "foqwsulbjiuslnue";
            var AccountPassword = "lfbeecgxvmwvzlna";
            string SenderEmail = "infiniumdevelopers@gmail.com";


            string to = ClientEmail;
            string from = SenderEmail;

            if (to.Length == 0)
            {
                MessageBox.Show("У клиента не указан Email. Отправка отчета невозможна");
                return;
            }

            to = to.Replace(';', ',');
            using (MailMessage message = new MailMessage(from, to))
            {
                message.Subject = "InfiniumAgent";
                message.Body = @"Здравствуйте. Отдел разработки программного обеспечения ООО «ОМЦ-ПРОФИЛЬ» предлагает Вам установить десктоп-приложение Infinium.Agent, которое является частью корпоративной системы управления ресурсами предприятия Infinium. В этой программе Вы сможете видеть состояние заказов Ваших клиентов на нашей фабрике в производстве, на складе и отгрузке. Кроме того, можете обмениваться сообщениями с нашими сотрудниками, читать и комментировать новости нашего предприятия. В программе также имеется актуальный каталог продукции. База данных находится на облачном хостинге, поэтому работает круглосуточно и вероятность сбоев сведена к минимуму. Скачать приложение можете по ссылке:
https://drive.google.com/file/d/0BzOa6U366p2pOFRESEdUN2hQSGc/view?usp=sharing&resourcekey=0-RqX0UQ9MBdrHuonIV88MRQ
Для входа в программу используйте логин " + Login + @" и пароль 1111. После авторизации рекомендуем сменить данные для входа.

--
С уважением, отдел разработки программного обеспечения ООО «ЗОВ - Профиль», Республика Беларусь, г. Гродно
infiniumdevelopers@gmail.com";
                //SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
                //{
                //    EnableSsl = true,
                //    UseDefaultCredentials = false,
                //    Credentials = new NetworkCredential(SenderEmail, AccountPassword)
                //};
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
                {
                    EnableSsl = true,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(SenderEmail, AccountPassword),
                    DeliveryMethod = SmtpDeliveryMethod.Network
                };
                try
                {
                    client.Send(message);
                    ClientEvents.AddEvent(ClientID, "Письмо с Агентом успешно отправлено", "MarketingClientsForm");
                }

                catch (ArgumentException ex)
                {
                    ClientEvents.AddEvent(ClientID, $"ArgumentException: {ex.Message}", "MarketingClientsForm");
                    MessageBox.Show("ArgumentException\r\n" + ex.Message);
                }
                catch (InvalidOperationException ex)
                {
                    ClientEvents.AddEvent(ClientID, $"InvalidOperationException: {ex.Message}", "MarketingClientsForm");
                    MessageBox.Show("InvalidOperationException\r\n" + ex.Message);
                }
                catch (SmtpException ex)
                {
                    ClientEvents.AddEvent(ClientID, $"SmtpException: {ex.Message}", "MarketingClientsForm");
                    MessageBox.Show("SmtpException\r\n" + ex.Message);
                }
                catch (Exception ex)
                {
                    ClientEvents.AddEvent(ClientID, $"Exception: {ex.Message}", "MarketingClientsForm");
                    MessageBox.Show("Exception\r\n" + ex.Message);
                }

                client.Dispose();
            }
        }

        private void FillClientRates()
        {
            var confirmDateTime = DateTime.Now;

            using (var da = new SqlDataAdapter(@"SELECT * FROM ClientRates WHERE CAST(Date AS Date) <= 
                    '" + confirmDateTime.ToString("yyyy-MM-dd") + "' ORDER BY clientId, Date DESC",
                       ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var dt = new DataTable())
                {
                    da.Fill(dt);
                    for (var i = 0; i < dt.Rows.Count; i++)
                    {
                        var clientId = Convert.ToInt32(dt.Rows[i]["clientId"]);
                        if (i > 0 && i < dt.Rows.Count)
                        {
                            if (clientId == Convert.ToInt32(dt.Rows[i - 1]["clientId"]))
                                continue;
                        }

                        if (dt.Rows[i]["USD"] == DBNull.Value || dt.Rows[i]["RUB"] == DBNull.Value ||
                            dt.Rows[i]["BYN"] == DBNull.Value)
                            continue;

                        var rows = ClientsDt.Select($"clientId={clientId}");
                        if (!rows.Any()) continue;

                        rows[0]["USD"] = Convert.ToDecimal(dt.Rows[i]["USD"]);
                        rows[0]["RUB"] = Convert.ToDecimal(dt.Rows[i]["RUB"]);
                        rows[0]["BYN"] = Convert.ToDecimal(dt.Rows[i]["BYN"]);
                    }
                }
            }
        }

        public void FillClientRatesNewYear()
        {
            var USD = 1.1m;
            var RUB = 100;
            var BYN = 3.35m;
            var date = new DateTime(2024, 01, 01);
            const string filter1 = @$"select TOP 0 * from ClientRates";

            const string filter2 = @$"SELECT  ClientID FROM  Clients
WHERE (CountryID = 1) AND (ClientID NOT IN (10774, 10771, 727, 725, 701, 759, 744, 737, 717, 708, 10777, 10776, 692, 724, 693, 684, 738, 695, 680, 766, 565, 447, 608, 698, 460, 552, 551, 327, 542))";

            var clientDt = new DataTable();
            using (var da = new SqlDataAdapter(filter2, ConnectionStrings.MarketingReferenceConnectionString))
            {
                da.Fill(clientDt);
            }
            using (var da = new SqlDataAdapter(filter1, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using var cb = new SqlCommandBuilder(da);
                using var dt = new DataTable();
                da.Fill(dt);

                for (var i = 0; i < clientDt.Rows.Count; i++)
                {
                    var clientId = Convert.ToInt32(clientDt.Rows[i]["clientId"]);

                    var newRow = dt.NewRow();
                    newRow["USD"] = USD;
                    newRow["RUB"] = RUB;
                    newRow["BYN"] = BYN;
                    newRow["date"] = date;
                    newRow["clientId"] = clientId;
                    dt.Rows.Add(newRow);
                }
                da.Update(dt);
            }
        }

        public void GetFixedPaymentRate(int ClientID, DateTime ConfirmDateTime, ref bool FixedPaymentRate, ref decimal USD, ref decimal RUB, ref decimal BYN)
        {
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM ClientRates WHERE CAST(Date AS Date) <= 
                    '" + ConfirmDateTime.ToString("yyyy-MM-dd") + "' AND ClientID = " + ClientID + " ORDER BY Date DESC",
                    ConnectionStrings.MarketingReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                    {
                        if (DT.Rows[0]["USD"] == DBNull.Value || DT.Rows[0]["RUB"] == DBNull.Value || DT.Rows[0]["BYN"] == DBNull.Value)
                            FixedPaymentRate = false;
                        else
                        {
                            FixedPaymentRate = true;
                            USD = Convert.ToDecimal(DT.Rows[0]["USD"]);
                            RUB = Convert.ToDecimal(DT.Rows[0]["RUB"]);
                            BYN = Convert.ToDecimal(DT.Rows[0]["BYN"]);
                        }
                    }
                    else
                        FixedPaymentRate = false;
                }
            }
            return;
        }

        public void GetDateRates(DateTime date, ref decimal EURBYRCurrency)
        {
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM DateRates WHERE CAST(Date AS Date) = 
                    '" + date.ToString("yyyy-MM-dd") + "'",
                    ConnectionStrings.MarketingReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                    {
                        EURBYRCurrency = Convert.ToDecimal(DT.Rows[0]["BYN"]);
                    }
                }
            }
            return;
        }
        
        public bool CBRDailyRates(DateTime date, ref decimal EURRUBCurrency, ref decimal USDRUBCurrency)
        {
            string EuroXML = "";
            string USDXml = "";

            string url = "http://www.cbr.ru/scripts/XML_daily.asp?date_req=" + date.ToString("dd/MM/yyyy");

            HttpWebRequest myHttpWebRequest;
            HttpWebResponse myHttpWebResponse;

            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                myHttpWebRequest.KeepAlive = false;
                myHttpWebRequest.AllowAutoRedirect = true;
                CookieContainer cookieContainer = new CookieContainer();
                myHttpWebRequest.CookieContainer = cookieContainer;
                myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
            }
            catch
            {
                return false;
            }

            XmlTextReader reader1 = new XmlTextReader(myHttpWebResponse.GetResponseStream());

            try
            {
                while (reader1.Read())
                {
                    switch (reader1.NodeType)
                    {
                        case XmlNodeType.Element:
                            if (reader1.Name == "Valute")
                            {
                                if (reader1.HasAttributes)
                                {
                                    while (reader1.MoveToNextAttribute())
                                    {
                                        if (reader1.Name == "ID")
                                        {
                                            if (reader1.Value == "R01239")
                                            {
                                                reader1.MoveToElement();
                                                EuroXML = reader1.ReadOuterXml();
                                            }
                                        }
                                        if (reader1.Name == "ID")
                                        {
                                            if (reader1.Value == "R01235")
                                            {
                                                reader1.MoveToElement();
                                                USDXml = reader1.ReadOuterXml();
                                            }
                                        }
                                    }
                                }
                            }

                            break;
                    }
                }
                XmlDocument euroXmlDocument = new XmlDocument();
                euroXmlDocument.LoadXml(EuroXML);
                XmlDocument usdXmlDocument = new XmlDocument();
                usdXmlDocument.LoadXml(USDXml);

                XmlNode xmlNode = euroXmlDocument.SelectSingleNode("Valute/Value");
                EURRUBCurrency = Convert.ToDecimal(xmlNode.InnerText = xmlNode.InnerText.Replace('.', ','));
                xmlNode = usdXmlDocument.SelectSingleNode("Valute/Value");
                USDRUBCurrency = Convert.ToDecimal(xmlNode.InnerText = xmlNode.InnerText.Replace('.', ','));
            }
            catch (WebException ex)
            {
                MessageBox.Show(ex.Message + " . КУРСЫ МОЖНО БУДЕТ ВВЕСТИ ВРУЧНУЮ");
                return false;
            }
            return true;
        }

        public void FixOrderEvent(int ClientID, string Event)
        {
            string SelectCommand = @"SELECT TOP 0 * FROM NewMegaOrdersEvents";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DataRow NewRow = DT.NewRow();
                        NewRow["MegaOrderID"] = ClientID;
                        NewRow["Event"] = Event;
                        NewRow["EventDate"] = Security.GetCurrentDate();
                        NewRow["EventUserID"] = Security.CurrentUserID;
                        DT.Rows.Add(NewRow);
                        DA.Update(DT);
                    }
                }
            }
        }

        public void RemoveClient(int ClientID)
        {
            using (var da = new SqlDataAdapter("DELETE FROM Clients" +
                " WHERE ClientID = " + ClientID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (DataTable dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                        ClientsDt.Clear();
                        ClientsDataAdapter.Fill(ClientsDt);
                        FillСlientsExcluziveCountries();
                    }
                }
            }
        }

        public void RemoveClientOrders(int ClientID)
        {
            using (var da = new SqlDataAdapter("DELETE FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders WHERE ClientID =" + ClientID + ")))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }

            using (var da = new SqlDataAdapter("DELETE FROM NewFrontsOrders WHERE MainOrderID IN (SELECT MainOrderID FROM NewMainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM NewMegaOrders WHERE ClientID =" + ClientID + "))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }
            using (var da = new SqlDataAdapter("DELETE FROM FrontsOrders WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders WHERE ClientID =" + ClientID + "))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }
            using (var da = new SqlDataAdapter("DELETE FROM NewDecorOrders WHERE MainOrderID IN (SELECT MainOrderID FROM NewMainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM NewMegaOrders WHERE ClientID =" + ClientID + "))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }
            using (var da = new SqlDataAdapter("DELETE FROM DecorOrders WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders WHERE ClientID =" + ClientID + "))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }

            using (var da = new SqlDataAdapter("DELETE FROM BatchDetails WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders WHERE ClientID =" + ClientID + "))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }
            using (var da = new SqlDataAdapter("DELETE FROM Packages WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders WHERE ClientID =" + ClientID + "))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }

            using (var da = new SqlDataAdapter("DELETE FROM NewMainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM NewMegaOrders WHERE ClientID =" + ClientID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }
            using (var da = new SqlDataAdapter("DELETE FROM MainOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders WHERE ClientID =" + ClientID + ")",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }

            using (var da = new SqlDataAdapter("DELETE FROM NewMegaOrders" +
                " WHERE ClientID = " + ClientID, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (DataTable dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }
            using (var da = new SqlDataAdapter("DELETE FROM MegaOrders" +
                " WHERE ClientID = " + ClientID, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (var dtTable = new DataTable())
                    {
                        da.Fill(dtTable);
                    }
                }
            }
        }

        private DataTable dtRolePermissions;

        public void GetPermissions(int UserID, string FormName)
        {
            if (dtRolePermissions == null)
                dtRolePermissions = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM UserRoles WHERE UserID = " + UserID +
                " AND RoleID IN (SELECT RoleID FROM Roles WHERE ModuleID IN " +
                " (SELECT ModuleID FROM Modules WHERE FormName = '" + FormName + "'))", ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(dtRolePermissions);
            }
        }

        public bool PermissionGranted(int RoleID)
        {
            DataRow[] Rows = dtRolePermissions.Select("RoleID = " + RoleID);
            return Rows.Count() > 0;
        }

        public void SaveClientsRatesFromExcel()
        {
            var table = new DataTable();
            table.Columns.Add(new DataColumn("ClientID", System.Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("USD", System.Type.GetType("System.Decimal")));
            table.Columns.Add(new DataColumn("RUB", System.Type.GetType("System.Decimal")));
            table.Columns.Add(new DataColumn("BYN", System.Type.GetType("System.Decimal")));
            table.TableName = "ImportedTable";

            var s = Clipboard.GetText();
            var lines = s.Split('\n');
            var data = new List<string>(lines);

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

            using var da = new SqlDataAdapter(@"SELECT TOP 0 * FROM ClientRates",
                ConnectionStrings.MarketingReferenceConnectionString);
            using (new SqlCommandBuilder(da))
            {
                using (var dt = new DataTable())
                {
                    da.Fill(dt);
                    for (var i = 0; i < table.Rows.Count; i++)
                    {
                        var clientId = Convert.ToInt32(table.Rows[i]["ClientID"]);
                        var USD = Convert.ToDecimal(table.Rows[i]["USD"]);
                        var RUB = Convert.ToDecimal(table.Rows[i]["RUB"]);
                        var BYN = Convert.ToDecimal(table.Rows[i]["BYN"]);
                        
                        var NewRow = dt.NewRow();
                        NewRow["clientId"] = clientId;
                        NewRow["Date"] = "2024-05-01 00:00:00.000";
                        NewRow["USD"] = USD;
                        NewRow["RUB"] = RUB;
                        NewRow["BYN"] = BYN;
                        dt.Rows.Add(NewRow);
                    }

                    da.Update(dt);
                }
            }
        }
    }

    public class ClientsToExcel
    {
        private HSSFWorkbook _workbook;

        private readonly Dictionary<string, string> _engToRusDictionary = new();

        public ClientsToExcel()
        {
            GetClients();
            PrepareColumns();
        }

        private DataTable _clientsDt;

        private void PrepareColumns()
        {
            _engToRusDictionary.Add("ClientID", "ID");
            _engToRusDictionary.Add("ClientName", "Клиент");
            _engToRusDictionary.Add("Manager", "Менеджер");
            _engToRusDictionary.Add("Country", "Страна");
            _engToRusDictionary.Add("City", "Город");
            _engToRusDictionary.Add("Email", "Email");
            _engToRusDictionary.Add("DelayOfPayment", "Отсрочка");
            _engToRusDictionary.Add("UNN", "UNN");
            _engToRusDictionary.Add("ClientGroupName", "Группа");
            _engToRusDictionary.Add("PriceGroup", "Ценовая группа");
            _engToRusDictionary.Add("Enabled", "Активен");
        }

        private void GetClients()
        {
            _clientsDt = new DataTable();

            const string selectCommand = @"SELECT Clients.ClientID, Clients.ClientName, ClientsManagers.Name as Manager, 
Countries.Name AS Country, Clients.City, Clients.Email, Clients.DelayOfPayment, Clients.UNN, 
ClientGroups.ClientGroupName, Clients.PriceGroup, Clients.Enabled FROM Clients AS Clients 
INNER JOIN ClientsManagers ON Clients.ManagerID = ClientsManagers.ManagerID 
INNER JOIN ClientGroups ON Clients.ClientGroupID = ClientGroups.ClientGroupID
INNER JOIN infiniu2_catalog.dbo.Countries AS Countries ON Clients.CountryID = Countries.CountryID order by Clients.ClientName";

            using var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingReferenceConnectionString);
            da.Fill(_clientsDt);
        }

        public void CreateReport(string fileName)
        {
            _workbook = new HSSFWorkbook();
            
            #region Create fonts and styles

            var clientNameFont = _workbook.CreateFont();
            clientNameFont.FontHeightInPoints = 14;
            clientNameFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            clientNameFont.FontName = "Calibri";

            var mainFont = _workbook.CreateFont();
            mainFont.FontHeightInPoints = 13;
            mainFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            mainFont.FontName = "Calibri";

            var clientNameStyle = _workbook.CreateCellStyle();
            clientNameStyle.SetFont(clientNameFont);

            var mainStyle = _workbook.CreateCellStyle();
            mainStyle.SetFont(mainFont);

            var headerFont = _workbook.CreateFont();
            headerFont.FontHeightInPoints = 13;
            headerFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            headerFont.FontName = "Calibri";

            var headerStyle = _workbook.CreateCellStyle();
            headerStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            headerStyle.BottomBorderColor = HSSFColor.BLACK.index;
            headerStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            headerStyle.LeftBorderColor = HSSFColor.BLACK.index;
            headerStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            headerStyle.RightBorderColor = HSSFColor.BLACK.index;
            headerStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            headerStyle.TopBorderColor = HSSFColor.BLACK.index;
            headerStyle.SetFont(headerFont);
            
            var simpleFont = _workbook.CreateFont();
            simpleFont.FontHeightInPoints = 12;
            simpleFont.FontName = "Calibri";

            var simpleCellStyle = _workbook.CreateCellStyle();
            simpleCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            simpleCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            simpleCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            simpleCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            simpleCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            simpleCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            simpleCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            simpleCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            simpleCellStyle.SetFont(simpleFont);

            var cellStyle = _workbook.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            cellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            cellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            cellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            cellStyle.RightBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            cellStyle.TopBorderColor = HSSFColor.BLACK.index;
            cellStyle.SetFont(simpleFont);
            
            #endregion
            
            Report(headerStyle, simpleCellStyle, cellStyle);

            var tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            var file = new FileInfo(tempFolder + @"\" + fileName + ".xls");
            var j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + fileName + "(" + j++ + ").xls");
            }

            var newFile = new FileStream(file.FullName, FileMode.Create);
            _workbook.Write(newFile);
            newFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
        }


        private void Report(HSSFCellStyle headerStyle, HSSFCellStyle simpleCellStyle, HSSFCellStyle cellStyle)
        {
            var rowIndex = 0;
            
            var sheet1 = _workbook.CreateSheet("Клиенты");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, .12);
            sheet1.SetMargin(HSSFSheet.RightMargin, .07);
            sheet1.SetMargin(HSSFSheet.TopMargin, .20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, .20);

            //sheet1.SetColumnWidth(0, 13 * 256);
            //sheet1.SetColumnWidth(1, 50 * 256);
            //sheet1.SetColumnWidth(2, 15 * 256);
            //sheet1.SetColumnWidth(3, 10 * 256);
            //sheet1.SetColumnWidth(4, 25 * 256);
            //sheet1.SetColumnWidth(5, 30 * 256);
            //sheet1.SetColumnWidth(6, 22 * 256);
            //sheet1.SetColumnWidth(7, 40 * 256);
            //sheet1.SetColumnWidth(8, 10 * 256);
            //sheet1.SetColumnWidth(9, 13 * 256);
            //sheet1.SetColumnWidth(10, 30 * 256);
            
            //string[] columns = {
            //"№ задания",
            //"Наименование объекта корпусной мебели",
            //"Облицовка",
            //"Платина",
            //"Цвет наполнителя",
            //"Примечание",
            //"Инвертарный номер",
            //"Бухгалтерское наименование детали",
            //"Кол-во",
            //"Стоимость",
            //"Дата принятия на склад",
            //};

            HSSFCell cell4;
            var i = 0;
            foreach (var column in _engToRusDictionary)
            {
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(rowIndex), i, column.Value);
                cell4.CellStyle = headerStyle;
                i++;
            }

            rowIndex++;
            
            for (var x = 0; x < _clientsDt.Rows.Count; x++)
            {
                for (var y = 0; y < _clientsDt.Columns.Count; y++)
                {
                    var t = _clientsDt.Rows[x][y].GetType();

                    switch (t.Name)
                    {
                        case "Decimal":
                        {
                            var cell = sheet1.CreateRow(rowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToDouble(_clientsDt.Rows[x][y]));

                            cell.CellStyle = cellStyle;
                            continue;
                        }
                        case "Int32":
                        {
                            var cell = sheet1.CreateRow(rowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToInt32(_clientsDt.Rows[x][y]));
                            cell.CellStyle = simpleCellStyle;
                            continue;
                        }
                        case "Int64":
                        {
                            var cell = sheet1.CreateRow(rowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToInt64(_clientsDt.Rows[x][y]));
                            cell.CellStyle = simpleCellStyle;
                            continue;
                        }
                        case "String":
                        case "Boolean":
                        case "DBNull":
                        {
                            var cell = sheet1.CreateRow(rowIndex).CreateCell(y);
                            cell.SetCellValue(_clientsDt.Rows[x][y].ToString());
                            cell.CellStyle = simpleCellStyle;
                            continue;
                        }
                        case "DateTime":
                        {
                            var cell = sheet1.CreateRow(rowIndex).CreateCell(y);
                            cell.SetCellValue(_clientsDt.Rows[x][y].ToString());
                            cell.CellStyle = simpleCellStyle;
                            continue;
                        }
                    }
                }
                rowIndex++;
            }
        }

    }
}
