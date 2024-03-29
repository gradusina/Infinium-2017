﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace Infinium
{
    public partial class MarketingFilterProductsForm : Form
    {
        public static bool NeedFilter = true;

        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;
        private int FactoryID = 1;

        private DataTable Fronts;
        private DataTable DecorProducts;

        private List<int> listBox1_selectionhistory;
        private List<int> listBox2_selectionhistory;

        private ArrayList FrontIDs;
        private ArrayList ProductIDs;

        private Form MainForm = null;

        public MarketingFilterProductsForm(int iFactoryID, ref ArrayList aFrontIDs, ref ArrayList aProductIDs)
        {
            InitializeComponent();

            FactoryID = iFactoryID;
            FactoryID = 1;
            FrontIDs = aFrontIDs;
            ProductIDs = aProductIDs;

            Initialize();
            CheckAllFrontsButton.Checked = true;
            CheckAllFrontsButton_Click(null, null);

            while (!SplashForm.bCreated) ;
        }

        private void Initialize()
        {
            listBox1_selectionhistory = new List<int>();
            listBox2_selectionhistory = new List<int>();

            Fronts = new DataTable();
            DecorProducts = new DataTable();

            string SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore 
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1 AND FactoryID = " + FactoryID + @" )
                ORDER BY TechStoreName";
            using (System.Data.SqlClient.SqlDataAdapter DA = new System.Data.SqlClient.SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(Fronts);
            }
            //using (System.Data.SqlClient.SqlDataAdapter DA = new System.Data.SqlClient.SqlDataAdapter(
            //    "SELECT FrontID, FrontName FROM Fronts WHERE FrontID IN (SELECT FrontID FROM infiniu2_catalog.dbo.FrontsConfig WHERE FactoryID = " + FactoryID + ") ORDER BY FrontName",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(Fronts);
            //}

            using (System.Data.SqlClient.SqlDataAdapter DA = new System.Data.SqlClient.SqlDataAdapter(
                "SELECT ProductID, ProductName, MeasureID, ReportParam FROM DecorProducts WHERE ProductID IN (SELECT ProductID FROM infiniu2_catalog.dbo.DecorConfig WHERE FactoryID = " + FactoryID + ") ORDER BY ProductName",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorProducts);
            }

            FrontsCheckedListBox.DataSource = Fronts;
            FrontsCheckedListBox.DisplayMember = "FrontName";
            FrontsCheckedListBox.ValueMember = "FrontID";

            DecorCheckedListBox.DataSource = DecorProducts;
            DecorCheckedListBox.DisplayMember = "ProductName";
            DecorCheckedListBox.ValueMember = "ProductID";
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
                        MainForm.Activate();
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
                        MainForm.Activate();
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

        private void SplitPackagesForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MarketingSplitPackagesForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void FrontsCheckedListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int actualcount = listBox1_selectionhistory.Count;
            if (actualcount == 1)
            {
                if (Control.ModifierKeys == Keys.Shift)
                {
                    int lastindex = listBox1_selectionhistory[0];
                    int currentindex = FrontsCheckedListBox.SelectedIndex;
                    int upper = Math.Max(lastindex, currentindex);
                    int lower = Math.Min(lastindex, currentindex);
                    for (int i = lower; i <= upper; i++)
                    {
                        FrontsCheckedListBox.SetItemCheckState(i, CheckState.Checked);
                    }
                }
                listBox1_selectionhistory.Clear();
                listBox1_selectionhistory.Add(FrontsCheckedListBox.SelectedIndex);
            }
            else
            {
                listBox1_selectionhistory.Clear();
                listBox1_selectionhistory.Add(FrontsCheckedListBox.SelectedIndex);
            }
        }

        private void DecorCheckedListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int actualcount = listBox2_selectionhistory.Count;
            if (actualcount == 1)
            {
                if (Control.ModifierKeys == Keys.Shift)
                {
                    int lastindex = listBox2_selectionhistory[0];
                    int currentindex = DecorCheckedListBox.SelectedIndex;
                    int upper = Math.Max(lastindex, currentindex);
                    int lower = Math.Min(lastindex, currentindex);
                    for (int i = lower; i <= upper; i++)
                    {
                        DecorCheckedListBox.SetItemCheckState(i, CheckState.Checked);
                    }
                }
                listBox2_selectionhistory.Clear();
                listBox2_selectionhistory.Add(DecorCheckedListBox.SelectedIndex);
            }
            else
            {
                listBox2_selectionhistory.Clear();
                listBox2_selectionhistory.Add(DecorCheckedListBox.SelectedIndex);
            }
        }

        private void CheckAllFrontsButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < FrontsCheckedListBox.Items.Count; i++)
                FrontsCheckedListBox.SetItemChecked(i, true);
        }

        private void UnCheckAllFrontsButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < FrontsCheckedListBox.Items.Count; i++)
                FrontsCheckedListBox.SetItemChecked(i, false);
        }

        private void CheckAllDecorButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < DecorCheckedListBox.Items.Count; i++)
                DecorCheckedListBox.SetItemChecked(i, true);
        }

        private void UnCheckAllDecorButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < DecorCheckedListBox.Items.Count; i++)
                DecorCheckedListBox.SetItemChecked(i, false);
        }

        private void GetFronts()
        {
            foreach (object itemChecked in FrontsCheckedListBox.CheckedItems)
            {
                DataRowView castedItem = (DataRowView)itemChecked;
                //int id = Convert.ToInt32(castedItem["FrontID"]);
                FrontIDs.Add(Convert.ToInt32(castedItem["FrontID"]));
            }
        }

        private void GetDecor()
        {
            foreach (object itemChecked in DecorCheckedListBox.CheckedItems)
            {
                DataRowView castedItem = (DataRowView)itemChecked;
                ProductIDs.Add(Convert.ToInt32(castedItem["ProductID"]));
            }
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            FrontIDs.Clear();
            ProductIDs.Clear();

            GetFronts();
            GetDecor();
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
            NeedFilter = true;
        }

        private void CancelPackButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
            NeedFilter = false;
        }

        private void CheckAllFrontsButton_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckAllFrontsButton.Checked)
                CheckAllFrontsButton_Click(null, null);
            else
                UnCheckAllFrontsButton_Click(null, null);
        }

        private void CheckAllDecorButton_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckAllDecorButton.Checked)
                CheckAllDecorButton_Click(null, null);
            else
                UnCheckAllDecorButton_Click(null, null);
        }
    }
}
