﻿using Infinium.Store;

using System;
using System.Drawing;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class PersonalInventoryForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private bool NeedSplash = false;
        private bool NeedCheckAll = true;
        private int CurrentFactoryID = 0;
        private int CurrentPersonID = 0;
        private int FormEvent = 0;
        private int Month = 0;
        private int Year = 0;

        private Form TopForm = null;
        private AddInventoryRestForm AddInventoryRestForm;
        private WriteOffParameters Parameters;

        private PersonalInventoryManager InventoryManager;

        public PersonalInventoryForm(int iCurrentFactoryID, int iCurrentPersonID,
            int iMonth, int iYear)
        {
            InitializeComponent();

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;
            CurrentFactoryID = iCurrentFactoryID;
            CurrentPersonID = iCurrentPersonID;
            Month = iMonth;
            Year = iYear;
            Initialize();
            label1.Text = "Infinium. Основной склад. Инвентаризация. " + new DateTime(Year, Month, 1).ToString("MMMM", CultureInfo.CurrentCulture);
            while (!SplashForm.bCreated) ;
        }

        private void InventoryForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;

            NeedSplash = true;
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                this.Opacity = 1;

                if (FormEvent == eClose || FormEvent == eHide)
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        this.Hide();
                    }

                    return;
                }

                if (FormEvent == eShow)
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                    return;
                }

            }

            if (FormEvent == eClose || FormEvent == eHide)
            {
                if (Convert.ToDecimal(this.Opacity) != Convert.ToDecimal(0.00))
                    this.Opacity = Convert.ToDouble(Convert.ToDecimal(this.Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        this.Hide();
                    }
                }

                return;
            }


            if (FormEvent == eShow || FormEvent == eShow)
            {
                if (this.Opacity != 1)
                    this.Opacity += 0.05;
                else
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }

                return;
            }
        }

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void Initialize()
        {
            InventoryManager = new PersonalInventoryManager(Month, Year)
            {
                CurrentFactoryID = CurrentFactoryID,
                CurrentPersonID = CurrentPersonID
            };
            InventoryManager.Initialize();
            //InventoryManager.GetGGG();
            Parameters = new WriteOffParameters();
            StoreDG.DataSource = InventoryManager.StoreList;

            SubGroupsDataGrid.DataSource = InventoryManager.SubGroupsList;
            GroupsDataGrid.DataSource = InventoryManager.GroupsList;
            StoreGridSettings();
            CheckStoreColumns(ref StoreDG);
        }

        private void checkboxHeader_CheckedChanged(object sender, EventArgs e)
        {
            if (!NeedCheckAll)
                return;
            for (int i = 0; i < StoreDG.RowCount; i++)
            {
                int StoreID = Convert.ToInt32(StoreDG["PersonalStoreID", i].Value);
                StoreDG["EditEnd", i].Value = ((ComponentFactory.Krypton.Toolkit.KryptonCheckBox)StoreDG.Controls.Find("checkboxHeader", true)[0]).Checked;
            }
            StoreDG.EndEdit();
        }

        private void AddCheckBoxColumn()
        {
            ComponentFactory.Krypton.Toolkit.KryptonCheckBox checkboxHeader = new ComponentFactory.Krypton.Toolkit.KryptonCheckBox()
            {
                PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black,
                Name = "checkboxHeader",
                Text = string.Empty,
                Size = new Size(18, 18),
                Location = new Point(19, 14)
            };
            checkboxHeader.CheckedChanged += new EventHandler(checkboxHeader_CheckedChanged);

            StoreDG.Controls.Add(checkboxHeader);
        }

        public void StoreGridSettings()
        {
            AddCheckBoxColumn();

            GroupsDataGrid.Columns["TechStoreGroupID"].Visible = false;

            SubGroupsDataGrid.Columns["TechStoreSubGroupID"].Visible = false;
            SubGroupsDataGrid.Columns["TechStoreGroupID"].Visible = false;
            SubGroupsDataGrid.Columns["Notes"].Visible = false;
            SubGroupsDataGrid.Columns["Notes1"].Visible = false;
            SubGroupsDataGrid.Columns["Notes2"].Visible = false;

            SubGroupsDataGrid.Columns["TechStoreSubGroupName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            SubGroupsDataGrid.Columns["TechStoreSubGroupName"].MinimumWidth = 100;

            StoreDG.Columns.Add(InventoryManager.ColorsColumn);
            StoreDG.Columns.Add(InventoryManager.PatinaColumn);
            StoreDG.Columns.Add(InventoryManager.CoversColumn);

            foreach (DataGridViewColumn Column in StoreDG.Columns)
            {
                Column.ReadOnly = true;
            }
            StoreDG.Columns["EditEnd"].ReadOnly = false;

            StoreDG.Columns["PersonalStoreID"].Visible = false;
            StoreDG.Columns["ColorID"].Visible = false;
            StoreDG.Columns["PatinaID"].Visible = false;
            StoreDG.Columns["StoreItemID"].Visible = false;
            StoreDG.Columns["CurrencyTypeID"].Visible = false;
            StoreDG.Columns["FactoryID"].Visible = false;
            StoreDG.Columns["MovementInvoiceID"].Visible = false;
            StoreDG.Columns["EditEnd"].Visible = false;
            StoreDG.Columns["InvoiceCount"].Visible = false;
            StoreDG.Columns["PriceEUR"].Visible = false;
            StoreDG.Columns["CurrentCount"].Visible = false;

            StoreDG.Columns["StoreItemColumn"].HeaderText = "Наименование";
            StoreDG.Columns["Length"].HeaderText = "Длина, мм";
            StoreDG.Columns["Width"].HeaderText = "Ширина, мм";
            StoreDG.Columns["Height"].HeaderText = "Высота, мм";
            StoreDG.Columns["Thickness"].HeaderText = "Толщина, мм";
            StoreDG.Columns["Diameter"].HeaderText = "Диаметр, мм";
            StoreDG.Columns["Admission"].HeaderText = "Допуск, мм";
            StoreDG.Columns["Weight"].HeaderText = "Вес, кг";
            StoreDG.Columns["Capacity"].HeaderText = "Емкость, л";
            StoreDG.Columns["Notes"].HeaderText = "Примечание";
            StoreDG.Columns["Price"].HeaderText = "Цена";
            StoreDG.Columns["VAT"].HeaderText = "НДС";
            StoreDG.Columns["Cost"].HeaderText = "Сумма";
            StoreDG.Columns["VATCost"].HeaderText = "Сумма с НДС";
            StoreDG.Columns["StartMonthCount"].HeaderText = "ОСТн";
            StoreDG.Columns["MonthInvoiceCount"].HeaderText = "Приход";
            StoreDG.Columns["ExpenseCount"].HeaderText = "Расход";
            StoreDG.Columns["SellingCount"].HeaderText = "Реализация";
            StoreDG.Columns["EndMonthCount"].HeaderText = "ОСТк";
            StoreDG.Columns["EditEnd"].HeaderText = string.Empty;
            StoreDG.Columns["InvoiceCount"].HeaderText = "Приход";
            StoreDG.Columns["InvNotes"].HeaderText = "Прим. инв-ции";
            StoreDG.Columns["PersonalStoreID"].HeaderText = "ID";
            StoreDG.Columns["CreateDateTime"].HeaderText = "Дата сохранения";

            StoreDG.AutoGenerateColumns = false;

            int DisplayIndex = 0;
            StoreDG.Columns["EditEnd"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["StoreItemColumn"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["CoversColumn"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["ColorsColumn"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["PatinaColumn"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Diameter"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Thickness"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Length"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Height"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Width"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Admission"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Capacity"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Weight"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["StartMonthCount"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["MonthInvoiceCount"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["ExpenseCount"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["SellingCount"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["EndMonthCount"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Price"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Cost"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["VAT"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["VATCost"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["Notes"].DisplayIndex = DisplayIndex++;
            StoreDG.Columns["PersonalStoreID"].DisplayIndex = DisplayIndex++;


            NumberFormatInfo nfi1 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = "",
                CurrencyDecimalDigits = 3,

                NumberGroupSeparator = " ",
                NumberDecimalDigits = 3,
                NumberDecimalSeparator = ","
            };
            StoreDG.Columns["Price"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Price"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["VAT"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["VAT"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["Cost"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Cost"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["VATCost"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["VATCost"].DefaultCellStyle.FormatProvider = nfi1;

            StoreDG.Columns["Thickness"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Thickness"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["Length"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Length"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["Height"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Height"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["Width"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Width"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["Admission"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Admission"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["Diameter"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Diameter"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["Capacity"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Capacity"].DefaultCellStyle.FormatProvider = nfi1;
            StoreDG.Columns["Weight"].DefaultCellStyle.Format = "N";
            StoreDG.Columns["Weight"].DefaultCellStyle.FormatProvider = nfi1;

            StoreDG.Columns["EditEnd"].Frozen = true;

            StoreDG.Columns["PersonalStoreID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            StoreDG.Columns["PersonalStoreID"].Width = 50;
            StoreDG.Columns["EditEnd"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            StoreDG.Columns["EditEnd"].Width = 50;
            StoreDG.Columns["StoreItemColumn"].MinimumWidth = 100;
            StoreDG.Columns["StoreItemColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["ColorsColumn"].MinimumWidth = 100;
            StoreDG.Columns["ColorsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["PatinaColumn"].MinimumWidth = 100;
            StoreDG.Columns["PatinaColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["CoversColumn"].MinimumWidth = 100;
            StoreDG.Columns["CoversColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            StoreDG.Columns["Thickness"].MinimumWidth = 60;
            StoreDG.Columns["Thickness"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Length"].MinimumWidth = 60;
            StoreDG.Columns["Length"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Thickness"].MinimumWidth = 60;
            StoreDG.Columns["Thickness"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Height"].MinimumWidth = 60;
            StoreDG.Columns["Height"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Width"].MinimumWidth = 60;
            StoreDG.Columns["Width"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Admission"].MinimumWidth = 60;
            StoreDG.Columns["Admission"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Diameter"].MinimumWidth = 60;
            StoreDG.Columns["Diameter"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Capacity"].MinimumWidth = 60;
            StoreDG.Columns["Capacity"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Weight"].MinimumWidth = 60;
            StoreDG.Columns["Weight"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Price"].MinimumWidth = 60;
            StoreDG.Columns["Price"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["MonthInvoiceCount"].MinimumWidth = 60;
            StoreDG.Columns["MonthInvoiceCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["StartMonthCount"].MinimumWidth = 60;
            StoreDG.Columns["StartMonthCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["ExpenseCount"].MinimumWidth = 60;
            StoreDG.Columns["ExpenseCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["SellingCount"].MinimumWidth = 60;
            StoreDG.Columns["SellingCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["EndMonthCount"].MinimumWidth = 60;
            StoreDG.Columns["EndMonthCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["CurrentCount"].MinimumWidth = 60;
            StoreDG.Columns["CurrentCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["VAT"].MinimumWidth = 60;
            StoreDG.Columns["VAT"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Cost"].MinimumWidth = 60;
            StoreDG.Columns["Cost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["VATCost"].MinimumWidth = 130;
            StoreDG.Columns["VATCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["Notes"].MinimumWidth = 60;
            StoreDG.Columns["Notes"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            StoreDG.Columns["InvNotes"].MinimumWidth = 60;
            StoreDG.Columns["InvNotes"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (TopForm != null)
                    TopForm.Activate();
            }
        }


        private void InvStoreDataGrid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (StoreDG.Columns[e.ColumnIndex].Name == "EditEnd")
                return;

            string InvNotes = string.Empty;
            int StoreID = 0;
            if (StoreDG.SelectedRows.Count != 0 && StoreDG.SelectedRows[0].Cells["InvNotes"].Value != DBNull.Value)
                InvNotes = StoreDG.SelectedRows[0].Cells["InvNotes"].Value.ToString();
            if (StoreDG.SelectedRows.Count != 0 && StoreDG.SelectedRows[0].Cells["PersonalStoreID"].Value != DBNull.Value)
                StoreID = Convert.ToInt32(StoreDG.SelectedRows[0].Cells["PersonalStoreID"].Value);

            decimal CurrentCount = InventoryManager.PlaningEndCount(StoreID);

            //if (CurrentCount < 1)
            //    return;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            AddInventoryRestForm = new AddInventoryRestForm(CurrentCount, InvNotes);

            TopForm = AddInventoryRestForm;
            AddInventoryRestForm.ShowDialog();
            TopForm = null;

            PhantomForm.Close();

            PhantomForm.Dispose();

            if (AddInventoryRestForm.IsOKPress)
            {
                int InventoryID = InventoryManager.CurrentInventoryID;
                int UserID = Security.CurrentUserID;
                decimal FactCount = AddInventoryRestForm.FactCount;
                string Notes = AddInventoryRestForm.Notes;

                AddInventoryRestForm.Dispose();
                AddInventoryRestForm = null;
                GC.Collect();

                InventoryManager.AddInventoryDetail(Security.GetCurrentDate(), InventoryID, StoreID, CurrentCount, FactCount, UserID, Notes);
                InventoryManager.ChangeCurrentFields(FactCount, Notes);
                InventoryManager.InventaryEndEdit(true);
                CheckStoreColumns(ref StoreDG);

            }
            else
            {
                AddInventoryRestForm.Dispose();
                AddInventoryRestForm = null;
                GC.Collect();
            }
        }

        private void CheckStoreColumns(ref PercentageDataGrid Grid)
        {
            foreach (DataGridViewColumn Column in Grid.Columns)
            {
                foreach (DataGridViewRow Row in Grid.Rows)
                {
                    if (Column.Name == "StartMonthCount"
                        || Column.Name == "EndMonthCount"
                        || Column.Name == "SellingCount"
                        || Column.Name == "InvoiceCount"
                        || Column.Name == "ExpenseCount"
                        || Column.Name == "MonthInvoiceCount")
                        continue;

                    if (Row.Cells[Column.Index].FormattedValue.ToString().Length > 0)
                    {
                        Grid.Columns[Column.Index].Visible = true;
                        break;
                    }
                    else
                        Grid.Columns[Column.Index].Visible = false;
                }
            }

            if (Grid.Columns.Contains("EditEnd"))
                Grid.Columns["EditEnd"].Visible = true;
            if (Grid.Columns.Contains("Producer"))
                Grid.Columns["Producer"].Visible = false;
            if (Grid.Columns.Contains("DecorAssignmentID"))
                Grid.Columns["DecorAssignmentID"].Visible = false;
            if (Grid.Columns.Contains("BestBefore"))
                Grid.Columns["BestBefore"].Visible = false;
            if (Grid.Columns.Contains("Produced"))
                Grid.Columns["Produced"].Visible = false;
            if (Grid.Columns.Contains("BatchNumber"))
                Grid.Columns["BatchNumber"].Visible = false;
            if (Grid.Columns.Contains("ManufacturerID"))
                Grid.Columns["ManufacturerID"].Visible = false;
            if (Grid.Columns.Contains("IsArrived"))
                Grid.Columns["IsArrived"].Visible = false;
            if (Grid.Columns.Contains("PersonalStoreID"))
                Grid.Columns["PersonalStoreID"].Visible = false;
            if (Grid.Columns.Contains("MeasureID"))
                Grid.Columns["MeasureID"].Visible = false;

            Grid.Columns["ColorID"].Visible = false;
            Grid.Columns["CoverID"].Visible = false;
            Grid.Columns["PatinaID"].Visible = false;
            Grid.Columns["StoreItemID"].Visible = false;
            Grid.Columns["CurrencyTypeID"].Visible = false;
            Grid.Columns["FactoryID"].Visible = false;
            Grid.Columns["MovementInvoiceID"].Visible = false;
            Grid.Columns["CurrentCount"].Visible = false;
            Grid.Columns["PriceEUR"].Visible = false;
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            int InventoryID = InventoryManager.CurrentInventoryID;
            InventoryManager.SaveInventoryDetails(InventoryID);
            InventoryManager.SaveStore();
            InventoryManager.RefreshStore();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            //FormEvent = eClose;
            //AnimateTimer.Enabled = true;
        }

        private void InvStoreDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if ((InvStoreDataGrid.Columns[e.ColumnIndex].Name == "EditEnd"))
            //{
            //    InventoryManager.InventaryEndEdit(true);
            //}
        }

        private void InvStoreDataGrid_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            if (InventoryManager.CheckFactCount(Convert.ToInt32(StoreDG.Rows[e.RowIndex].Cells["PersonalStoreID"].Value)))
            {
                StoreDG.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                StoreDG.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
            }
            else
            {
                if (Convert.ToBoolean(StoreDG.Rows[e.RowIndex].Cells["EditEnd"].Value))
                {
                    StoreDG.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(31, 158, 0);
                    StoreDG.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                }
                else
                {
                    StoreDG.Rows[e.RowIndex].DefaultCellStyle.BackColor = Security.GridsBackColor;
                    StoreDG.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                }
            }

        }

        public void AllEditEnd()
        {
            int AllCheckedCount = 0;
            for (int i = 0; i < StoreDG.RowCount; i++)
            {
                if (Convert.ToBoolean(StoreDG["EditEnd", i].Value))
                    AllCheckedCount++;
            }
            if (AllCheckedCount == StoreDG.RowCount)
                ((ComponentFactory.Krypton.Toolkit.KryptonCheckBox)StoreDG.Controls.Find("checkboxHeader", true)[0]).Checked = true;
            else
                ((ComponentFactory.Krypton.Toolkit.KryptonCheckBox)StoreDG.Controls.Find("checkboxHeader", true)[0]).Checked = false;
        }

        private void SubGroupsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (InventoryManager == null)
                return;

            int TechStoreSubGroupID = 0;
            if (SubGroupsDataGrid.SelectedRows.Count != 0 && SubGroupsDataGrid.SelectedRows[0].Cells["TechStoreSubGroupID"].Value != DBNull.Value)
                TechStoreSubGroupID = Convert.ToInt32(SubGroupsDataGrid.SelectedRows[0].Cells["TechStoreSubGroupID"].Value);

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление данных.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;
                InventoryManager.FilterStore(TechStoreSubGroupID);
                NeedCheckAll = false;
                //AllEditEnd();
                NeedCheckAll = true;
                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                InventoryManager.FilterStore(TechStoreSubGroupID);
                NeedCheckAll = false;
                //AllEditEnd();
                NeedCheckAll = true;
            }
            CheckStoreColumns(ref StoreDG);
        }

        private void AddExcessButton_Click(object sender, EventArgs e)
        {

        }

        private void PrintInventoryButton_Click(object sender, EventArgs e)
        {
            string InventoryName = "Инвентаризационная опись, ОМЦ-ПРОФИЛЬ, " + DateTimeFormatInfo.CurrentInfo.GetMonthName(Month) + " " + Year.ToString();
            if (CurrentFactoryID == 2)
                InventoryName = "Инвентаризационная опись, ЗОВ-ТПС, " + DateTimeFormatInfo.CurrentInfo.GetMonthName(Month) + " " + Year.ToString();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            PersonalStoreInventoryList InventoryList = new PersonalStoreInventoryList(CurrentFactoryID, InventoryName, Month, Year);
            InventoryList.InventoryReport2(InventoryName);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void AddInventoryButton_Click(object sender, EventArgs e)
        {
            if (InventoryManager == null)
                return;

            if (!InventoryManager.InventoryExist(Month, Year))
            {
                //Infinium.LightMessageBox.Show(ref TopForm, false,
                //      "На этот месяц инвентаризационная опись уже создана",
                //      "Ошибка");
                //return;
                InventoryManager.CreateMonthInventory();
            }

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание новой\r\nинвентаризации. Подождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            InventoryManager.CreateMonthInventoryDetails();
            InventoryManager.RefreshStore();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void AddExcessButton_Click_1(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;

            NewArrivalForm NewArrivalForm = new NewArrivalForm();

            TopForm = NewArrivalForm;

            NewArrivalForm.ShowDialog();

            NewArrivalForm.Close();
            NewArrivalForm.Dispose();

            TopForm = null;

            InventoryManager.UpdateTables();
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            WriteOffParametersMenu MovementReportMenu = new WriteOffParametersMenu(this, ref Parameters);

            TopForm = MovementReportMenu;
            MovementReportMenu.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();

            MovementReportMenu.Dispose();
            TopForm = null;

            if (!Parameters.OKPress)
                return;

            InventoryManager.InventoryReport2("Акт списания", Parameters.WriteOffDate, Parameters.User1, Parameters.User2, Parameters.User3, Parameters.User4, Parameters.User5);
        }

        private void GroupsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (InventoryManager == null)
                return;

            int TechStoreGroupID = 0;
            if (GroupsDataGrid.SelectedRows.Count != 0 && GroupsDataGrid.SelectedRows[0].Cells["TechStoreGroupID"].Value != DBNull.Value)
                TechStoreGroupID = Convert.ToInt32(GroupsDataGrid.SelectedRows[0].Cells["TechStoreGroupID"].Value);

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление данных.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;
                InventoryManager.FilterStoreSubGroups(TechStoreGroupID);
                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
                InventoryManager.FilterStoreSubGroups(TechStoreGroupID);
        }
    }
}
