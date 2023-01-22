using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace Infinium.Modules.ZOV.Clients
{
    public class Clients
    {
        private PercentageDataGrid ClientsDataGrid = null;
        private PercentageDataGrid ClientsGroupsDataGrid = null;
        private PercentageDataGrid ManagersDataGrid = null;
        private PercentageDataGrid ContactsDataGrid = null;
        private PercentageDataGrid ShopAddressesDataGrid = null;

        private DataTable ClientsGroupsDataTable = null;
        public DataTable AllClientsDataTable = null;
        private DataTable ClientsDataTable = null;
        private DataTable ManagersDataTable = null;

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
            ClientsDataGrid = tClientsDataGrid;
            ClientsGroupsDataGrid = tClientsGroupsDataGrid;
            ManagersDataGrid = tManagersDataGrid;
            ContactsDataGrid = tContactsDataGrid;
            ShopAddressesDataGrid = tShopAddressesDataGrid;

            Initialize();
        }

        private void Create()
        {
            AllClientsDataTable = new DataTable();
            ClientsGroupsDataTable = new DataTable();
            ClientsDataTable = new DataTable();
            ManagersDataTable = new DataTable();
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
            ManagersDataAdapter.Fill(ManagersDataTable);

            ClientsGroupsDataAdapter = new SqlDataAdapter("SELECT * FROM ClientsGroups ORDER BY ClientGroupName",
                ConnectionStrings.ZOVReferenceConnectionString);
            ClientsGroupsCommandBuilder = new SqlCommandBuilder(ClientsGroupsDataAdapter);
            ClientsGroupsDataAdapter.Fill(ClientsGroupsDataTable);
            
            ClientsDataAdapter = new SqlDataAdapter("SELECT * FROM Clients ORDER BY ClientName",
                ConnectionStrings.ZOVReferenceConnectionString);
            ClientsCommandBuilder = new SqlCommandBuilder(ClientsDataAdapter);
            ClientsDataAdapter.Fill(ClientsDataTable);

            using (SqlDataAdapter da = new SqlDataAdapter(@"SELECT Clients.ClientName, ClientsGroups.ClientGroupName, Managers.Name
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
            ClientsGroupsBindingSource.DataSource = ClientsGroupsDataTable;
            ClientsBindingSource.DataSource = ClientsDataTable;
            ManagersBindingSource.DataSource = ManagersDataTable;
            ContactsBindingSource.DataSource = ContactsDataTable;
            ShopAddressesBindingSource.DataSource = ShopAddressesDataTable;

            ManagersDataGrid.DataSource = ManagersBindingSource;
            ClientsGroupsDataGrid.DataSource = ClientsGroupsBindingSource;
            ClientsDataGrid.DataSource = ClientsBindingSource;
            ContactsDataGrid.DataSource = ContactsBindingSource;
            ShopAddressesDataGrid.DataSource = ShopAddressesBindingSource;

            ClientsBindingSourceDisplayMember = "ClientName";

            ClientsBindingSourceValueMember = "ClientID";
        }

        private void SetClientsDataGrid()
        {
            ClientsDataGrid.Columns.Add(ClientsGroupsColumn);
            ClientsDataGrid.Columns.Add(ManagerColumn);

            foreach (DataGridViewColumn Column in ClientsDataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            ManagersDataGrid.Columns["Name"].HeaderText = "Менеджер";

            ClientsGroupsDataGrid.Columns["ClientGroupName"].HeaderText = "Название";

            ClientsDataGrid.Columns["ClientGroupID"].Visible = false;
            ClientsDataGrid.Columns["ManagerID"].Visible = false;
            ClientsDataGrid.Columns["Contacts"].Visible = false;
            ClientsDataGrid.Columns["MoveOk"].Visible = false;

            ClientsDataGrid.Columns["ClientID"].HeaderText = "ID";
            ClientsDataGrid.Columns["ClientName"].HeaderText = "Название организации";

            ClientsDataGrid.AutoGenerateColumns = false;

            ClientsDataGrid.Columns["ClientID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ClientsDataGrid.Columns["ClientID"].Width = 50;
            //ClientsDataGrid.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //ClientsDataGrid.Columns["ClientName"].MinimumWidth = 150;
            //ClientsDataGrid.Columns["ManagerColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //ClientsDataGrid.Columns["ManagerColumn"].MinimumWidth = 150;
            //ClientsDataGrid.Columns["ClientsGroupsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //ClientsDataGrid.Columns["ClientsGroupsColumn"].MinimumWidth = 150;

            int DisplayIndex = 0;
            ClientsDataGrid.Columns["ClientID"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["ClientName"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["ClientsGroupsColumn"].DisplayIndex = DisplayIndex++;
            ClientsDataGrid.Columns["ManagerColumn"].DisplayIndex = DisplayIndex++;
        }

        private DataGridViewComboBoxColumn ManagerColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "ManagerColumn",
                    HeaderText = "Менеджер",
                    DataPropertyName = "ManagerID",

                    DataSource = new DataView(ManagersDataTable),
                    ValueMember = "ManagerID",
                    DisplayMember = "Name",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    ReadOnly = false
                };
                return Column;
            }
        }

        private DataGridViewComboBoxColumn ClientsGroupsColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "ClientsGroupsColumn",
                    HeaderText = "Группа клиентов",
                    DataPropertyName = "ClientGroupID",

                    DataSource = new DataView(ClientsGroupsDataTable),
                    ValueMember = "ClientGroupID",
                    DisplayMember = "ClientGroupName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    ReadOnly = false
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
            ShopAddressesDataGrid.Columns["Address"].HeaderText = "Адреса салонов";

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

            DataRow[] Rows = ClientsDataTable.Select("ClientID = " + ClientID);

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
            tShopAddressesDataGrid.Columns["Address"].HeaderText = "Адреса салонов";

            tShopAddressesDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            //tShopAddressesDataGrid.Columns["Address"].DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;
            //tShopAddressesDataGrid.Columns["Address"].DefaultCellStyle.Font = new System.Drawing.Font("SEGOE UI", 13.0F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Pixel, ((byte)(0)));
        }

        public void SaveClientsGroups()
        {
            ClientsGroupsDataAdapter.Update(ClientsGroupsDataTable);
            ClientsGroupsDataTable.Clear();
            ClientsGroupsDataAdapter.Fill(ClientsGroupsDataTable);
        }

        public void SaveAllClients()
        {
            ClientsDataAdapter.Update(ClientsDataTable);
            ClientsDataTable.Clear();
            ClientsDataAdapter.Fill(ClientsDataTable);
        }

        public void SaveManagers()
        {
            ManagersDataAdapter.Update(ManagersDataTable);
            ManagersDataTable.Clear();
            ManagersDataAdapter.Fill(ManagersDataTable);
        }

        public void AddClient(string Name, int ClientGroupID, int ManagerID)
        {
            string Contacts = GetContactsXML(NewContactsDataTable);

            DataRow Row = ClientsDataTable.NewRow();
            Row["ManagerID"] = ManagerID;
            Row["ClientName"] = Name;
            Row["ClientGroupID"] = ClientGroupID;
            Row["Contacts"] = Contacts;

            ClientsDataTable.Rows.Add(Row);
            ClientsDataAdapter.Update(ClientsDataTable);
            ClientsDataTable.Clear();
            ClientsDataAdapter.Fill(ClientsDataTable);
        }

        public void EditClient(ref string ClientsName, ref int ClientGroupID, ref int ManagerID)
        {
            ClientsName = ((DataRowView)ClientsBindingSource.Current)["ClientName"].ToString();
            ClientGroupID = Convert.ToInt32(((DataRowView)ClientsBindingSource.Current)["ClientGroupID"]);
            ManagerID = Convert.ToInt32(((DataRowView)ClientsBindingSource.Current)["ManagerID"]);

            NewContactsDataTable.Clear();
            NewContactsDataTable = ContactsDataTable.Copy();
            NewContactsBindingSource.DataSource = NewContactsDataTable;

            NewShopAddressesDataTable.Clear();
            NewShopAddressesDataTable = ShopAddressesDataTable.Copy();
            NewShopAddressesBindingSource.DataSource = NewShopAddressesDataTable;
        }

        public void SaveClient(string Name, int ClientGroupID, int ManagerID, int ClientID)
        {
            string Contacts = GetContactsXML(NewContactsDataTable);

            int index = ClientsBindingSource.Position;

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM Clients WHERE ClientID = " + ClientID + " ORDER BY ClientName",
                ConnectionStrings.ZOVReferenceConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);
                        if (DT.Rows.Count > 0)
                        {
                            DT.Rows[0]["ClientName"] = Name;
                            DT.Rows[0]["ClientGroupID"] = ClientGroupID;

                            DT.Rows[0]["Contacts"] = Contacts;
                            DT.Rows[0]["ManagerID"] = ManagerID;

                            DA.Update(DT);
                            ClientsDataTable.Clear();
                            ClientsDataAdapter.Fill(ClientsDataTable);
                            ClientsBindingSource.Position = ClientsBindingSource.Find("ClientID", ClientID);
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
    }

    public class ZOVClientsToExcel
    {
        private int pos01 = 0;

        private HSSFWorkbook hssfworkbook;

        private HSSFFont fConfirm;
        private HSSFFont fHeader;
        private HSSFFont fColumnName;
        private HSSFFont fMainContent;
        private HSSFFont fTotalInfo;

        private HSSFCellStyle csConfirm;
        private HSSFCellStyle csHeader;
        private HSSFCellStyle csColumnName;
        private HSSFCellStyle csMainContent;
        private HSSFCellStyle csTotalInfo;

        public ZOVClientsToExcel()
        {
            hssfworkbook = new HSSFWorkbook();
            CreateFonts();
            CreateCellStyles();
        }

        private void CreateFonts()
        {
            fConfirm = hssfworkbook.CreateFont();
            fConfirm.FontHeightInPoints = 12;
            fConfirm.FontName = "Calibri";

            fHeader = hssfworkbook.CreateFont();
            fHeader.FontHeightInPoints = 12;
            fHeader.Boldweight = 12 * 256;
            fHeader.FontName = "Calibri";

            fColumnName = hssfworkbook.CreateFont();
            fColumnName.FontHeightInPoints = 12;
            fColumnName.Boldweight = 12 * 256;
            fColumnName.FontName = "Calibri";

            fMainContent = hssfworkbook.CreateFont();
            fMainContent.FontHeightInPoints = 11;
            fMainContent.FontName = "Calibri";

            fTotalInfo = hssfworkbook.CreateFont();
            fTotalInfo.FontHeightInPoints = 11;
            fTotalInfo.Boldweight = 11 * 256;
            fTotalInfo.FontName = "Calibri";
        }

        private void CreateCellStyles()
        {
            csConfirm = hssfworkbook.CreateCellStyle();
            csConfirm.SetFont(fConfirm);

            csHeader = hssfworkbook.CreateCellStyle();
            csHeader.SetFont(fHeader);

            csColumnName = hssfworkbook.CreateCellStyle();
            csColumnName.BorderBottom = HSSFCellStyle.BORDER_THIN;
            csColumnName.BottomBorderColor = HSSFColor.BLACK.index;
            csColumnName.BorderLeft = HSSFCellStyle.BORDER_THIN;
            csColumnName.LeftBorderColor = HSSFColor.BLACK.index;
            csColumnName.BorderRight = HSSFCellStyle.BORDER_THIN;
            csColumnName.RightBorderColor = HSSFColor.BLACK.index;
            csColumnName.BorderTop = HSSFCellStyle.BORDER_THIN;
            csColumnName.TopBorderColor = HSSFColor.BLACK.index;
            csColumnName.Alignment = HSSFCellStyle.ALIGN_CENTER;
            csColumnName.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            csColumnName.WrapText = true;
            csColumnName.SetFont(fColumnName);

            csMainContent = hssfworkbook.CreateCellStyle();
            csMainContent.BorderBottom = HSSFCellStyle.BORDER_THIN;
            csMainContent.BottomBorderColor = HSSFColor.BLACK.index;
            csMainContent.BorderLeft = HSSFCellStyle.BORDER_THIN;
            csMainContent.LeftBorderColor = HSSFColor.BLACK.index;
            csMainContent.BorderRight = HSSFCellStyle.BORDER_THIN;
            csMainContent.RightBorderColor = HSSFColor.BLACK.index;
            csMainContent.BorderTop = HSSFCellStyle.BORDER_THIN;
            csMainContent.TopBorderColor = HSSFColor.BLACK.index;
            csMainContent.VerticalAlignment = HSSFCellStyle.ALIGN_RIGHT;
            csMainContent.SetFont(fMainContent);

            csTotalInfo = hssfworkbook.CreateCellStyle();
            csTotalInfo.BorderBottom = HSSFCellStyle.BORDER_THIN;
            csTotalInfo.BottomBorderColor = HSSFColor.BLACK.index;
            csTotalInfo.BorderLeft = HSSFCellStyle.BORDER_THIN;
            csTotalInfo.LeftBorderColor = HSSFColor.BLACK.index;
            csTotalInfo.BorderRight = HSSFCellStyle.BORDER_THIN;
            csTotalInfo.RightBorderColor = HSSFColor.BLACK.index;
            csTotalInfo.BorderTop = HSSFCellStyle.BORDER_THIN;
            csTotalInfo.TopBorderColor = HSSFColor.BLACK.index;
            csTotalInfo.Alignment = HSSFCellStyle.ALIGN_LEFT;
            csTotalInfo.SetFont(fTotalInfo);
        }

        public void ClearReport()
        {
            hssfworkbook = new HSSFWorkbook();
            CreateFonts();
            CreateCellStyles();
        }

        public void Export(DataTable table1)
        {
            HSSFSheet sheet01 = hssfworkbook.CreateSheet("Клиенты");
            sheet01.PrintSetup.PaperSize = (short)PaperSizeType.A4;
            sheet01.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet01.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet01.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet01.SetMargin(HSSFSheet.BottomMargin, (double).20);
            pos01 = 0;

            int displayIndex = 0;
            {
                HSSFCell cell1 = null;
                for (int i = 0; i < table1.Columns.Count; i++)
                {
                    sheet01.SetColumnWidth(i, 15 * 256);
                    string colName = table1.Columns[i].ToString();

                    cell1 = sheet01.CreateRow(pos01).CreateCell(displayIndex++);
                    cell1.SetCellValue(colName);
                    cell1.CellStyle = csColumnName;
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

                pos01++;

                //Содержимое таблицы
                for (int x = 0; x < table1.Rows.Count; x++)
                {
                    for (int y = 0; y < table1.Columns.Count; y++)
                    {
                        Type t = table1.Rows[x][y].GetType();

                        switch (t.Name)
                        {
                            case "Int32":
                                cell1 = sheet01.CreateRow(pos01).CreateCell(y);
                                cell1.SetCellValue(Convert.ToInt32(table1.Rows[x][y]));
                                cell1.CellStyle = csMainContent;
                                continue;
                            case "String":
                            case "DBNull":
                                cell1 = sheet01.CreateRow(pos01).CreateCell(y);
                                cell1.SetCellValue(table1.Rows[x][y].ToString());
                                cell1.CellStyle = csMainContent;
                                continue;
                        }
                    }

                    pos01++;
                }
            }

            pos01++;
            pos01++;
        }

        public void SaveFile(string FileName, bool bOpenFile)
        {
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
            ClearReport();

            if (bOpenFile)
                System.Diagnostics.Process.Start(file.FullName);
        }
    }
}
