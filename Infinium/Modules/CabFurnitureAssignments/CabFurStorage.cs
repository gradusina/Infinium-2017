﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Infinium.Modules.CabFurnitureAssignments
{
    class CabFurStorage
    {
        DataTable workShopsDt = null;
        DataTable racksDt = null;
        DataTable cellsDt = null;

        public BindingSource workShopsBs = null;
        public BindingSource racksBs = null;
        public BindingSource cellsBs = null;

        SqlDataAdapter workShopsDa = null;
        SqlCommandBuilder workShopsCb = null;

        SqlDataAdapter racksDa = null;
        SqlCommandBuilder racksCb = null;

        SqlDataAdapter cellsDa = null;
        SqlCommandBuilder cellsCb = null;

        public CabFurStorage()
        {
            CreateObjects();
            //ПОПРОБОВАТЬ НАОБОРОТ ВЫЗВАТЬ ФУНКЦИИ
            FillTables();
            BindingTables();
        }

        private void CreateObjects()
        {
            workShopsDt = new DataTable();
            racksDt = new DataTable();
            cellsDt = new DataTable();

            workShopsBs = new BindingSource();
            racksBs = new BindingSource();
            cellsBs = new BindingSource();
        }

        private void BindingTables()
        {
            workShopsBs = new BindingSource()
            {
                DataSource = workShopsDt
            };
            racksBs = new BindingSource()
            {
                DataSource = racksDt
            };
            cellsBs = new BindingSource()
            {
                DataSource = cellsDt
            };
        }

        private void FillTables()
        {
            string selectCommand = @"SELECT * FROM WorkShops ORDER BY Name";
            workShopsDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            workShopsCb = new SqlCommandBuilder(workShopsDa);
            workShopsDa.Fill(workShopsDt);

            selectCommand = @"SELECT * FROM Racks ORDER BY Name";
            racksDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            racksCb = new SqlCommandBuilder(racksDa);
            racksDa.Fill(racksDt);

            selectCommand = @"SELECT * FROM Cells ORDER BY Name";
            cellsDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            cellsCb = new SqlCommandBuilder(cellsDa);
            cellsDa.Fill(cellsDt);
        }

        #region WorkShops

        public bool HasWorkShops
        {
            get { return workShopsBs.Count > 0; }
        }

        public string CurrentWorkShopName
        {
            get
            {
                string name = string.Empty;
                if (workShopsBs.Current != null)
                    name = ((DataRowView)workShopsBs.Current).Row["Name"].ToString();
                return name;
            }
        }

        public int CurrentWorkShopId
        {
            get
            {
                int id = -1;
                if (workShopsBs.Current != null)
                    id = Convert.ToInt32(((DataRowView)workShopsBs.Current).Row["WorkShopID"]);
                return id;
            }
        }

        public int MaxWorkShopId
        {
            get { return Convert.ToInt32(workShopsDt.Compute("max([WorkShopID])", string.Empty)); }
        }

        public void SetWorkShopPosition(int Id)
        {
            workShopsBs.Position = workShopsBs.Find("WorkShopID", Id);
        }

        public void AddWorkShop(string name)
        {
            DataRow NewRow = workShopsDt.NewRow();
            NewRow["Name"] = name;
            workShopsDt.Rows.Add(NewRow);
        }

        public void EditWorkShop(int id, string name)
        {
            DataRow[] rows = workShopsDt.Select("WorkShopID = " + id);
            if (rows.Count() > 0)
                rows[0]["Name"] = name;
        }

        public void RemoveWorkShop(int id)
        {
            foreach (DataRow row in workShopsDt.Select("WorkShopID = " + id))
            {
                RemoveRackByWorkshop(id);
                row.Delete();
            }
        }

        public void UpdateWorkShops()
        {
            workShopsDt.Clear();
            string selectCommand = @"SELECT * FROM WorkShops ORDER BY Name";
            workShopsDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            workShopsCb = new SqlCommandBuilder(workShopsDa);
            workShopsDa.Fill(workShopsDt);
        }

        public void SaveWorkShops()
        {
            workShopsDa.Update(workShopsDt);
        }

        public void SearchWorkShops(int TechStoreID, int CoverID, int PatinaID, int InsetColorID)
        {
            workShopsDt.Clear();
            string selectCommand = "SELECT * FROM WorkShops WHERE WorkShopID IN (SELECT WorkShopID FROM Racks WHERE RackID IN (SELECT RackID FROM Cells WHERE CellID IN (" +
                " SELECT CellID FROM CabFurniturePackages" +
                " WHERE TechStoreID=" + TechStoreID +
                " AND CoverID=" + CoverID +
                " AND PatinaID=" + PatinaID +
                " AND InsetColorID=" + InsetColorID + "))) ORDER BY Name";
            workShopsDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            workShopsDa.Fill(workShopsDt);
        }
        #endregion

        #region Racks

        public bool HasRacks
        {
            get { return racksBs.Count > 0; }
        }

        public string CurrentRackName
        {
            get
            {
                string name = string.Empty;
                if (racksBs.Current != null)
                    name = ((DataRowView)racksBs.Current).Row["Name"].ToString();
                return name;
            }
        }

        public int CurrentRackId
        {
            get
            {
                int id = -1;
                if (racksBs.Current != null)
                    id = Convert.ToInt32(((DataRowView)racksBs.Current).Row["RackID"]);
                return id;
            }
        }

        public int MaxRackId
        {
            get { return Convert.ToInt32(racksDt.Compute("max([RackID])", string.Empty)); }
        }

        public void SetRackPosition(int Id)
        {
            racksBs.Position = racksBs.Find("RackID", Id);
        }

        public void FilterRacksByWorkShop(int workShopId)
        {
            racksBs.Filter = "WorkShopID=" + workShopId;
            racksBs.MoveFirst();
        }

        public void AddRack(string name, int workShopId)
        {
            DataRow NewRow = racksDt.NewRow();
            NewRow["Name"] = name;
            NewRow["WorkShopID"] = workShopId;
            racksDt.Rows.Add(NewRow);
        }

        public void EditRack(int id, string name)
        {
            DataRow[] rows = racksDt.Select("RackID = " + id);
            if (rows.Count() > 0)
                rows[0]["Name"] = name;
        }

        public void RemoveRack(int id)
        {
            foreach (DataRow row in racksDt.Select("RackID = " + id))
            {
                RemoveCellByRack(id);
                row.Delete();
            }
        }
        
        private void RemoveRackByWorkshop(int workShopID)
        {
            foreach (DataRow row in racksDt.Select("WorkShopID = " + workShopID))
            {
                int id = Convert.ToInt32(row["RackId"]);
                RemoveCellByRack(id);
                row.Delete();
            }
        }

        public void UpdateRacks()
        {
            racksDt.Clear();

            string selectCommand = @"SELECT * FROM Racks ORDER BY Name";
            racksDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            racksCb = new SqlCommandBuilder(racksDa);
            racksDa.Fill(racksDt);

            racksBs.MoveFirst();
        }

        public void SaveRacks()
        {
            racksDa.Update(racksDt);
        }

        public void SearchRacks(int TechStoreID, int CoverID, int PatinaID, int InsetColorID)
        {
            racksDt.Clear();
            string selectCommand = "SELECT * FROM Racks WHERE RackID IN (SELECT RackID FROM Cells WHERE CellID IN (" +
                " SELECT CellID FROM CabFurniturePackages" +
                " WHERE TechStoreID=" + TechStoreID +
                " AND CoverID=" + CoverID +
                " AND PatinaID=" + PatinaID +
                " AND InsetColorID=" + InsetColorID + ")) ORDER BY Name";
            racksDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            racksDa.Fill(racksDt);
        }
        #endregion

        private void UnbindPackagesFromCell(int cellId)
        {
            string filter = @"SELECT CabFurniturePackageID, CellID, 
AddToStorage, AddToStorageDateTime, AddToStorageUserID, 
BindToCellUserID, BindToCellDateTime, QualityControlInUserID, QualityControlInDateTime, 
QualityControlOutUserID, QualityControlOutDateTime, QualityControl FROM CabFurniturePackages 
WHERE CellID=" + cellId;

            using (SqlDataAdapter DA = new SqlDataAdapter(filter, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) <= 0) return;

                        for (int i = 0; i < DT.Rows.Count; i++)
                        {
                            DT.Rows[i]["CellID"] = -1;
                            DT.Rows[i]["BindToCellUserID"] = DBNull.Value;
                            DT.Rows[i]["BindToCellDateTime"] = DBNull.Value;
                            DT.Rows[i]["AddToStorage"] = 0;
                            DT.Rows[i]["AddToStorageDateTime"] = DBNull.Value;
                            DT.Rows[i]["AddToStorageUserID"] = DBNull.Value;
                            DT.Rows[i]["QualityControlInDateTime"] = DBNull.Value;
                            DT.Rows[i]["QualityControlInUserID"] = DBNull.Value;
                            DT.Rows[i]["QualityControlOutDateTime"] = DBNull.Value;
                            DT.Rows[i]["QualityControlOutUserID"] = DBNull.Value;
                            DT.Rows[i]["QualityControl"] = -1;
                        }
                        DA.Update(DT);
                    }
                }
            }
        }
        
        private void UnbindPackagesFromCell(int[] cellsIds)
        {
            string filter = string.Empty;
            foreach (int item in cellsIds)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = @"SELECT CabFurniturePackageID, CellID, 
AddToStorage, AddToStorageDateTime, AddToStorageUserID, 
BindToCellUserID, BindToCellDateTime, QualityControlInUserID, QualityControlInDateTime, 
QualityControlOutUserID, QualityControlOutDateTime, QualityControl FROM CabFurniturePackages 
WHERE CellID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(filter, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) <= 0) return;

                        for (int i = 0; i < DT.Rows.Count; i++)
                        {
                            DT.Rows[i]["CellID"] = -1;
                            DT.Rows[i]["BindToCellUserID"] = DBNull.Value;
                            DT.Rows[i]["BindToCellDateTime"] = DBNull.Value;
                            DT.Rows[i]["AddToStorage"] = 0;
                            DT.Rows[i]["AddToStorageDateTime"] = DBNull.Value;
                            DT.Rows[i]["AddToStorageUserID"] = DBNull.Value;
                            DT.Rows[i]["QualityControlInDateTime"] = DBNull.Value;
                            DT.Rows[i]["QualityControlInUserID"] = DBNull.Value;
                            DT.Rows[i]["QualityControlOutDateTime"] = DBNull.Value;
                            DT.Rows[i]["QualityControlOutUserID"] = DBNull.Value;
                            DT.Rows[i]["QualityControl"] = -1;
                        }
                        DA.Update(DT);
                    }
                }
            }
        }

        #region Cells

        public bool HasCells
        {
            get { return cellsBs.Count > 0; }
        }

        public string CurrentCellName
        {
            get
            {
                string name = string.Empty;
                if (cellsBs.Current != null)
                    name = ((DataRowView)cellsBs.Current).Row["Name"].ToString();
                return name;
            }
        }

        public int CurrentCellId
        {
            get
            {
                int id = -1;
                if (cellsBs.Current != null)
                    id = Convert.ToInt32(((DataRowView)cellsBs.Current).Row["CellID"]);
                return id;
            }
        }

        public int MaxCellId
        {
            get { return Convert.ToInt32(cellsDt.Compute("max([CellID])", string.Empty)); }
        }

        public void FilterCellsByRack(int rackId)
        {
            cellsBs.Filter = "RackID=" + rackId;
            cellsBs.MoveFirst();
        }

        public void AddCell(string name, int rackId)
        {
            DataRow NewRow = cellsDt.NewRow();
            NewRow["Name"] = name;
            NewRow["RackID"] = rackId;
            cellsDt.Rows.Add(NewRow);
        }

        public void EditCell(int id, string name)
        {
            DataRow[] rows = cellsDt.Select("CellID = " + id);
            if (rows.Count() > 0)
                rows[0]["Name"] = name;
        }

        public void RemoveCell(int id)
        {
            foreach (DataRow row in cellsDt.Select("CellID = " + id))
            {
                UnbindPackagesFromCell(id);
                row.Delete();
            }
        }
        
        private void RemoveCellByRack(int rackId)
        {
            int[] cellsIds = new int[cellsDt.Select("RackId = " + rackId).Length];
            int i = 0;
            foreach (DataRow row in cellsDt.Select("RackId = " + rackId))
            {
                cellsIds[i++] = Convert.ToInt32(row["CellID"]);
                row.Delete();
            }
            UnbindPackagesFromCell(cellsIds);
        }

        public void UpdateCells()
        {
            cellsDt.Clear();

            string selectCommand = @"SELECT * FROM Cells ORDER BY Name";
            cellsDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            cellsCb = new SqlCommandBuilder(cellsDa);
            cellsDa.Fill(cellsDt);
        }

        public void SaveCells()
        {
            cellsDa.Update(cellsDt);
        }

        public bool SearchCells(int TechStoreID, int CoverID, int PatinaID, int InsetColorID)
        {
            cellsDt.Clear();
            string selectCommand = "SELECT * FROM Cells WHERE CellID IN (" +
                " SELECT CellID FROM CabFurniturePackages" +
                " WHERE TechStoreID=" + TechStoreID +
                " AND CoverID=" + CoverID +
                " AND PatinaID=" + PatinaID +
                " AND InsetColorID=" + InsetColorID + ") ORDER BY Name";
            cellsDa = new SqlDataAdapter(selectCommand, ConnectionStrings.StorageConnectionString);
            cellsDa.Fill(cellsDt);
            return cellsDt.Rows.Count > 0;
        }

        public List<CellLabelInfo> CreateCellLabels(int[] cells)
        {
            string filter = string.Empty;
            foreach (int item in cells)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = " WHERE Cells.CellID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            List<CellLabelInfo> Labels = new List<CellLabelInfo>();
            string SelectCommand = @"SELECT Cells.*, Racks.Name AS RackName, WorkShops.Name AS WorkShopName FROM Cells
                INNER JOIN Racks ON Cells.RackID=Racks.RackID
                INNER JOIN WorkShops ON Racks.WorkShopID=WorkShops.WorkShopID" + filter;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        for (int i = 0; i < DT.Rows.Count; i++)
                        {
                            string WorkShopName = DT.Rows[i]["WorkShopName"].ToString();
                            string RackName = DT.Rows[i]["RackName"].ToString();
                            string CellName = DT.Rows[i]["Name"].ToString();
                            int CellID = Convert.ToInt32(DT.Rows[i]["CellID"]);
                            string BarcodeNumber = GetBarcodeNumber(23, Convert.ToInt32(DT.Rows[i]["CellID"]));

                            CellLabelInfo LabelInfo = new CellLabelInfo();
                            LabelInfo.WorkShopName = WorkShopName;
                            LabelInfo.RackName = RackName;
                            LabelInfo.CellName = CellName;
                            LabelInfo.CellID = CellID;
                            LabelInfo.BarcodeNumber = BarcodeNumber;
                            LabelInfo.FactoryType = 1;

                            Labels.Add(LabelInfo);
                        }
                    }
                }
            }

            return Labels;
        }

        public string GetBarcodeNumber(int BarcodeType, int PackNumber)
        {
            string Type = "";
            if (BarcodeType.ToString().Length == 1)
                Type = "00" + BarcodeType.ToString();
            if (BarcodeType.ToString().Length == 2)
                Type = "0" + BarcodeType.ToString();
            if (BarcodeType.ToString().Length == 3)
                Type = BarcodeType.ToString();

            string Number = "";
            if (PackNumber.ToString().Length == 1)
                Number = "00000000" + PackNumber.ToString();
            if (PackNumber.ToString().Length == 2)
                Number = "0000000" + PackNumber.ToString();
            if (PackNumber.ToString().Length == 3)
                Number = "000000" + PackNumber.ToString();
            if (PackNumber.ToString().Length == 4)
                Number = "00000" + PackNumber.ToString();
            if (PackNumber.ToString().Length == 5)
                Number = "0000" + PackNumber.ToString();
            if (PackNumber.ToString().Length == 6)
                Number = "000" + PackNumber.ToString();
            if (PackNumber.ToString().Length == 7)
                Number = "00" + PackNumber.ToString();
            if (PackNumber.ToString().Length == 8)
                Number = "0" + PackNumber.ToString();

            StringBuilder BarcodeNumber = new StringBuilder(Type);
            BarcodeNumber.Append(Number);

            return BarcodeNumber.ToString();
        }

        #endregion
    }

    public struct CellLabelInfo
    {
        public string WorkShopName;
        public string RackName;
        public string CellName;
        public int CellID;
        public string BarcodeNumber;
        public int FactoryType;
    }

    public class CellLabel
    {
        Barcode Barcode;
        public PrintDocument PD;

        public int PaperHeight = 488;
        public int PaperWidth = 394;

        public int CurrentLabelNumber = 0;

        public int PrintedCount = 0;

        public bool Printed = false;

        SolidBrush FontBrush;

        Font ClientFont;
        Font DocFont;
        Font InfoFont;

        Pen Pen;

        Image ZTTPS;
        Image ZTProfil;
        Image STB;
        Image RST;

        public ArrayList LabelInfo;

        public CellLabel()
        {

            Barcode = new Barcode();

            InitializeFonts();
            InitializePrinter();

            ZTTPS = new Bitmap(Properties.Resources.ZOVTPS);
            ZTProfil = new Bitmap(Properties.Resources.ZOVPROFIL);
            STB = new Bitmap(Properties.Resources.eac);
            RST = new Bitmap(Properties.Resources.RST);

            LabelInfo = new System.Collections.ArrayList();
        }

        private void InitializePrinter()
        {
            PD = new PrintDocument();
            PD.DefaultPageSettings.PaperSize = new PaperSize("Label", PaperWidth, PaperHeight);
            PD.DefaultPageSettings.Landscape = true;
            PD.DefaultPageSettings.Margins.Bottom = 0;
            PD.DefaultPageSettings.Margins.Top = 0;
            PD.DefaultPageSettings.Margins.Left = 0;
            PD.DefaultPageSettings.Margins.Right = 0;
        }

        private void InitializeFonts()
        {
            FontBrush = new System.Drawing.SolidBrush(Color.Black);
            ClientFont = new Font("Arial", 20, FontStyle.Regular);
            DocFont = new Font("Arial", 14, FontStyle.Regular);
            InfoFont = new Font("Arial", 5.0f, FontStyle.Regular);
            Pen = new Pen(FontBrush)
            {
                Width = 1
            };
        }

        public void ClearLabelInfo()
        {
            CurrentLabelNumber = 0;
            PrintedCount = 0;
            LabelInfo.Clear();
            GC.Collect();
        }

        public void AddLabelInfo(ref CellLabelInfo LabelInfoItem)
        {
            LabelInfo.Add(LabelInfoItem);
        }

        public void PrintPage(object sender, PrintPageEventArgs ev)
        {
            if (PrintedCount >= LabelInfo.Count)
                return;
            else
                PrintedCount++;

            ev.Graphics.Clear(Color.White);

            ev.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

            int HorizLineClientBot = 18;
            ev.Graphics.DrawLine(Pen, 11, HorizLineClientBot, 467, HorizLineClientBot);

            ev.Graphics.DrawString("Цех: " + ((CellLabelInfo)LabelInfo[CurrentLabelNumber]).WorkShopName, ClientFont, FontBrush, 19, HorizLineClientBot + 82);

            ev.Graphics.DrawString("Стеллаж: " + ((CellLabelInfo)LabelInfo[CurrentLabelNumber]).RackName, ClientFont, FontBrush, 19, HorizLineClientBot + 142);

            ev.Graphics.DrawString("Ячейка: " + ((CellLabelInfo)LabelInfo[CurrentLabelNumber]).CellName, ClientFont, FontBrush, 19, HorizLineClientBot + 202);

            int HorizLinOrderBot = 75;
            int HorizLinSmallBarcode = 338;
            ev.Graphics.DrawLine(Pen, HorizLinSmallBarcode, HorizLinOrderBot, 467, HorizLinOrderBot);

            ev.Graphics.DrawString("№" + ((CellLabelInfo)LabelInfo[CurrentLabelNumber]).CellID.ToString(), DocFont, FontBrush, HorizLinSmallBarcode + 4, HorizLineClientBot + 4);
            ev.Graphics.DrawLine(Pen, HorizLinSmallBarcode, HorizLineClientBot, HorizLinSmallBarcode, HorizLinOrderBot);

            ev.Graphics.DrawImage(Barcode.GetBarcode(Barcode.BarcodeLength.Medium, 46, ((CellLabelInfo)LabelInfo[CurrentLabelNumber]).BarcodeNumber), 10, 317);

            Barcode.DrawBarcodeText(Barcode.BarcodeLength.Medium, ev.Graphics, ((CellLabelInfo)LabelInfo[CurrentLabelNumber]).BarcodeNumber, 9, 366);

            ev.Graphics.DrawImage(Barcode.GetBarcode(Barcode.BarcodeLength.Short, 15, ((CellLabelInfo)LabelInfo[CurrentLabelNumber]).BarcodeNumber), HorizLinSmallBarcode + 2, 54, 130, 15);

            if (((CellLabelInfo)LabelInfo[CurrentLabelNumber]).FactoryType == 2)
            {
                ev.Graphics.DrawImage(ZTTPS, 249, 320, 37, 45);
            }
            else
            {
                ev.Graphics.DrawImage(ZTProfil, 249, 320, 37, 45);
            }

            ev.Graphics.DrawImage(STB, 418, 319, 39, 27);
            //ev.Graphics.DrawImage(RST, 423, 357, 34, 27);

            ev.Graphics.DrawLine(Pen, 11, 315, 467, 315);
            ev.Graphics.DrawLine(Pen, 235, 315, 235, 385);

            ev.Graphics.DrawString("ГОСТ 16371-2014", InfoFont, FontBrush, 305, 318);

            if (((CellLabelInfo)LabelInfo[CurrentLabelNumber]).FactoryType == 2)
                ev.Graphics.DrawString("ООО \"ЗОВ-ТПС\"", InfoFont, FontBrush, 305, 328);
            else
                ev.Graphics.DrawString("ООО \"ОМЦ-ПРОФИЛЬ\"", InfoFont, FontBrush, 305, 328);

            ev.Graphics.DrawString("Республика Беларусь, 230011", InfoFont, FontBrush, 305, 338);
            ev.Graphics.DrawString("г. Гродно, ул. Герасимовича, 1", InfoFont, FontBrush, 305, 348);
            ev.Graphics.DrawString("тел\\факс: (0152) 52-14-70", InfoFont, FontBrush, 305, 358);
            //ev.Graphics.DrawString("Изготовлено: " + ((PackageLabelInfo)LabelInfo[CurrentLabelNumber]).AddToStorageDateTime, InfoFont, FontBrush, 305, 368);
            ev.Graphics.DrawString("Распечатено: " + Security.GetCurrentDate().ToString(), InfoFont, FontBrush, 305, 378);
            //ev.Graphics.DrawString(((PackageLabelInfo)LabelInfo[CurrentLabelNumber]).DispatchDate, DispatchFont, FontBrush, 237, 374);

            if (CurrentLabelNumber == LabelInfo.Count - 1 || PrintedCount >= LabelInfo.Count)
                ev.HasMorePages = false;

            if (CurrentLabelNumber < LabelInfo.Count - 1 && PrintedCount < LabelInfo.Count)
            {
                ev.HasMorePages = true;
                CurrentLabelNumber++;
            }
        }

        public void Print()
        {
            PD.DefaultPageSettings.PaperSize = new PaperSize("Label", PaperWidth + 40, PaperHeight);
            PD.DefaultPageSettings.Margins.Bottom = 0;
            PD.DefaultPageSettings.Margins.Top = 0;
            PD.DefaultPageSettings.Margins.Right = 0;
            PD.DefaultPageSettings.Margins.Left = 0;

            if (!Printed)
            {
                Printed = true;
                PD.PrintPage += new PrintPageEventHandler(PrintPage);
            }

            PD.Print();
        }
    }

    public class StorePackagesManager
    {
        DataTable PackageLabelsDT = null;
        DataTable BindPackageLabelsDT = null;
        DataTable QualityControlDT = null;
        DataTable AllQualityControlDT = null;
        //DataTable ExcessInvPackageLabelsDT = null;
        DataTable MissInvPackageLabelsDT = null;
        DataTable InvPackageLabelsDT = null;
        public BindingSource PackageLabelsBS = null;
        public BindingSource BindPackageLabelsBS = null;
        public BindingSource QualityControlBS = null;
        public BindingSource AllQualityControlBS = null;
        public BindingSource ExcessInvPackageLabelsBS = null;
        public BindingSource MissInvPackageLabelsBS = null;
        public BindingSource InvPackageLabelsBS = null;
        SqlDataAdapter PackageLabelsDA;
        SqlDataAdapter BindPackageLabelsDA;
        SqlDataAdapter QualityControlDA;
        SqlDataAdapter AllQualityControlDA;
        SqlDataAdapter InvPackageLabelsDA;

        DataTable PackageDetailsDT = null;
        public BindingSource PackageDetailsBS = null;
        SqlDataAdapter PackageDetailsDA;

        //DataTable ExcessInvPackageDetailsDT = null;
        public BindingSource ExcessInvPackageDetailsBS = null;

        DataTable MissInvPackageDetailsDT = null;
        public BindingSource MissInvPackageDetailsBS = null;

        DataTable InvPackageDetailsDT = null;
        public BindingSource InvPackageDetailsBS = null;
        SqlDataAdapter InvPackageDetailsDA;

        public StorePackagesManager()
        {
            PackageLabelsDT = new DataTable();
            BindPackageLabelsDT = new DataTable();
            QualityControlDT = new DataTable();
            AllQualityControlDT = new DataTable();
            //ExcessInvPackageLabelsDT = new DataTable();
            MissInvPackageLabelsDT = new DataTable();
            InvPackageLabelsDT = new DataTable();
            PackageDetailsDT = new DataTable();
            //ExcessInvPackageDetailsDT = new DataTable();
            MissInvPackageDetailsDT = new DataTable();
            InvPackageDetailsDT = new DataTable();

            PackageLabelsBS = new BindingSource();
            BindPackageLabelsBS = new BindingSource();
            QualityControlBS = new BindingSource();
            AllQualityControlBS = new BindingSource();
            ExcessInvPackageLabelsBS = new BindingSource();
            MissInvPackageLabelsBS = new BindingSource();
            InvPackageLabelsBS = new BindingSource();
            PackageDetailsBS = new BindingSource();
            ExcessInvPackageDetailsBS = new BindingSource();
            MissInvPackageDetailsBS = new BindingSource();
            InvPackageDetailsBS = new BindingSource();

            string SelectCommand = @"SELECT TOP 0 CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, Cells.Name FROM CabFurniturePackages 
                INNER JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID";
            PackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            PackageLabelsDA.Fill(PackageLabelsDT);
            PackageLabelsDT.Columns.Add(new DataColumn("Index", Type.GetType("System.Int32")));

            SelectCommand = @"SELECT TOP 0 CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name FROM CabFurniturePackages 
                INNER JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID";
            BindPackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            BindPackageLabelsDA.Fill(BindPackageLabelsDT);
            
            SelectCommand = @"SELECT TOP 0 CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name FROM CabFurniturePackages 
                INNER JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID";
            QualityControlDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            QualityControlDA.Fill(QualityControlDT);
                        
            SelectCommand = @"SELECT TOP 0 CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name FROM CabFurniturePackages 
                INNER JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID";
            AllQualityControlDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            AllQualityControlDA.Fill(AllQualityControlDT);

            SelectCommand = @"SELECT TOP 0 CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name as CellName, Racks.Name as RackName, Workshops.Name as WorkshopName, Workshops.WorkshopID FROM CabFurniturePackages 
                INNER JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID 
                INNER JOIN Racks ON Cells.RackID=Racks.RackID 
                INNER JOIN Workshops ON Racks.WorkshopID=Workshops.WorkshopID";
            InvPackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            //InvPackageLabelsDA.Fill(ExcessInvPackageLabelsDT);
            InvPackageLabelsDA.Fill(MissInvPackageLabelsDT);
            InvPackageLabelsDA.Fill(InvPackageLabelsDT);

            SelectCommand = @"SELECT TOP 0 C.PackNumber, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID, 
                CabFurniturePackageDetails.* FROM CabFurniturePackageDetails 
                INNER JOIN CabFurniturePackages AS C ON CabFurniturePackageDetails.CabFurniturePackageID=C.CabFurniturePackageID";
            PackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            PackageDetailsDA.Fill(PackageDetailsDT);

            SelectCommand = @"SELECT TOP 0 C.PackNumber, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID, 
                CabFurniturePackageDetails.* FROM CabFurniturePackageDetails 
                INNER JOIN CabFurniturePackages AS C ON CabFurniturePackageDetails.CabFurniturePackageID=C.CabFurniturePackageID";
            InvPackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            //InvPackageDetailsDA.Fill(ExcessInvPackageDetailsDT);
            InvPackageDetailsDA.Fill(MissInvPackageDetailsDT);
            InvPackageDetailsDA.Fill(InvPackageDetailsDT);

            PackageLabelsBS.DataSource = PackageLabelsDT;
            BindPackageLabelsBS.DataSource = BindPackageLabelsDT;
            QualityControlBS.DataSource = QualityControlDT;
            AllQualityControlBS.DataSource = AllQualityControlDT;
            ExcessInvPackageLabelsBS.DataSource = new DataView(InvPackageLabelsDT);
            MissInvPackageLabelsBS.DataSource = MissInvPackageLabelsDT;
            InvPackageLabelsBS.DataSource = new DataView(InvPackageLabelsDT);
            PackageDetailsBS.DataSource = PackageDetailsDT;
            ExcessInvPackageDetailsBS.DataSource = new DataView(InvPackageDetailsDT);
            MissInvPackageDetailsBS.DataSource = MissInvPackageDetailsDT;
            InvPackageDetailsBS.DataSource = new DataView(InvPackageDetailsDT);
        }

        public void GetPackagesLabels(int CellID)
        {
            PackageDetailsDT.Clear();
            string SelectCommand = @"SELECT C.PackNumber, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID,
                CabFurniturePackageDetails.* FROM CabFurniturePackageDetails 
                INNER JOIN CabFurniturePackages AS C ON CabFurniturePackageDetails.CabFurniturePackageID=C.CabFurniturePackageID
                WHERE CellID=" + CellID;
            PackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            PackageDetailsDA.Fill(PackageDetailsDT);

            PackageLabelsDT.Clear();
            SelectCommand = @"SELECT CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name FROM CabFurniturePackages 
                LEFT JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID WHERE CabFurniturePackages.CellID=" + CellID;
            PackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            PackageLabelsDA.Fill(PackageLabelsDT);

            for (int i = 0; i < PackageLabelsDT.Rows.Count; i++)
            {
                PackageLabelsDT.Rows[i]["Index"] = i + 1;
            }
        }

        public bool GetBindPackagesLabels(int CabFurniturePackageID)
        {
            BindPackageLabelsDT.Clear();
            string SelectCommand = @"SELECT CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name FROM CabFurniturePackages 
                LEFT JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID WHERE CabFurniturePackageID=" + CabFurniturePackageID;
            BindPackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            BindPackageLabelsDA.Fill(BindPackageLabelsDT);
            
            return BindPackageLabelsDT.Rows.Count > 0;
        }

        public bool GetQualityControl(int CabFurniturePackageID)
        {
            QualityControlDT.Clear();

            string SelectCommand = $@"SELECT CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name FROM CabFurniturePackages 
                LEFT JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID WHERE CabFurniturePackageID={CabFurniturePackageID}";
            QualityControlDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            QualityControlDA.Fill(QualityControlDT);
            
            return QualityControlDT.Rows.Count > 0;
        }
        
        public bool GetAllQualityControl()
        {
            AllQualityControlDT.Clear();

            string SelectCommand = @"SELECT CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name FROM CabFurniturePackages 
                LEFT JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID WHERE QualityControl=0";
            AllQualityControlDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            AllQualityControlDA.Fill(AllQualityControlDT);
            
            return AllQualityControlDT.Rows.Count > 0;
        }

        public void GetExcessInvPackagesLabels(int workshopID)
        {
            ExcessInvPackageLabelsBS.Filter = "WorkShopID IS NULL OR WorkShopID<>" + workshopID;
        }

        public bool GetMissInvPackagesLabels(int workshopID)
        {
            MissInvPackageLabelsDT.Clear();
            MissInvPackageDetailsDT.Clear();
            string[] results = new string[InvPackageLabelsDT.Rows.Count];
            for (int i = 0; i < InvPackageLabelsDT.Rows.Count; i++)
                results[i] = InvPackageLabelsDT.Rows[i]["CabFurniturePackageID"].ToString();
            if (results.Count() == 0)
                return false;

            string filter = string.Empty;
            foreach (string item in results)
                filter += item + ",";
            if (filter.Length > 0)
                filter = " NOT IN (" + filter.Substring(0, filter.Length - 1) + ") AND Workshops.WorkShopID=" + workshopID;

            string SelectCommand = @"SELECT C.PackNumber, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID,
                CabFurniturePackageDetails.* FROM CabFurniturePackageDetails 
                INNER JOIN CabFurniturePackages AS C ON CabFurniturePackageDetails.CabFurniturePackageID=C.CabFurniturePackageID
                LEFT JOIN Cells ON C.CellID=Cells.CellID 
                LEFT JOIN Racks ON Cells.RackID=Racks.RackID 
                LEFT JOIN Workshops ON Racks.WorkshopID=Workshops.WorkshopID WHERE CabFurniturePackageDetails.CabFurniturePackageID" + filter;
            InvPackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            InvPackageDetailsDA.Fill(MissInvPackageDetailsDT);

            SelectCommand = @"SELECT CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name as CellName, Racks.Name as RackName, Workshops.Name as WorkshopName, Workshops.WorkshopID FROM CabFurniturePackages 
                LEFT JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID 
                LEFT JOIN Racks ON Cells.RackID=Racks.RackID 
                LEFT JOIN Workshops ON Racks.WorkshopID=Workshops.WorkshopID WHERE CabFurniturePackages.CabFurniturePackageID" + filter;
            InvPackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            InvPackageLabelsDA.Fill(MissInvPackageLabelsDT);

            return MissInvPackageLabelsDT.Rows.Count > 0;
        }

        public bool GetInvPackagesLabels(int CabFurniturePackageID)
        {
            string SelectCommand = @"SELECT C.PackNumber, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID,
                CabFurniturePackageDetails.* FROM CabFurniturePackageDetails 
                INNER JOIN CabFurniturePackages AS C ON CabFurniturePackageDetails.CabFurniturePackageID=C.CabFurniturePackageID
                WHERE CabFurniturePackageDetails.CabFurniturePackageID=" + CabFurniturePackageID;
            InvPackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            InvPackageDetailsDA.Fill(InvPackageDetailsDT);

            DataTable dt = InvPackageLabelsDT.Clone();
            SelectCommand = @"SELECT CabFurniturePackageID, PackNumber, AddToStorageDateTime, RemoveFromStorageDateTime, QualityControlInDateTime, QualityControlOutDateTime, QualityControl, CabFurniturePackages.CellID, Cells.Name as CellName, Racks.Name as RackName, Workshops.Name as WorkshopName, Workshops.WorkshopID FROM CabFurniturePackages 
                LEFT JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID 
                LEFT JOIN Racks ON Cells.RackID=Racks.RackID 
                LEFT JOIN Workshops ON Racks.WorkshopID=Workshops.WorkshopID WHERE CabFurniturePackages.CabFurniturePackageID=" + CabFurniturePackageID;
            InvPackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            InvPackageLabelsDA.Fill(dt);

            foreach (DataRow dr in dt.Rows)
                InvPackageLabelsDT.Rows.Add(dr.ItemArray);
            dt.Dispose();
            return InvPackageLabelsDT.Rows.Count > 0;
        }

        public void FilterPackagesDetails(int CabFurniturePackageID)
        {
            PackageDetailsBS.Filter = "CabFurniturePackageID =" + CabFurniturePackageID;
            PackageDetailsBS.MoveFirst();
        }

        public void FilterMissInvPackagesDetails(int CabFurniturePackageID)
        {
            MissInvPackageDetailsBS.Filter = "CabFurniturePackageID =" + CabFurniturePackageID;
            MissInvPackageDetailsBS.MoveFirst();
        }

        public void FilterInvPackagesDetails(int CabFurniturePackageID)
        {
            InvPackageDetailsBS.Filter = "CabFurniturePackageID =" + CabFurniturePackageID;
            InvPackageDetailsBS.MoveFirst();
        }

        public void ClearInvTables()
        {
            //ExcessInvPackageLabelsDT.Clear();
            MissInvPackageLabelsDT.Clear();
            InvPackageLabelsDT.Clear();
            //ExcessInvPackageDetailsDT.Clear();
            MissInvPackageDetailsDT.Clear();
            InvPackageDetailsDT.Clear();
        }

        public bool IsPackageScan(int CabFurniturePackageID)
        {
            DataRow[] rows = InvPackageLabelsDT.Select("CabFurniturePackageID=" + CabFurniturePackageID);
            return rows.Count() > 0;
        }

        public bool IsCellExist(int cellId)
        {
            bool bExist = false;
            string SelectCommand = @"SELECT CellID FROM Cells WHERE CellID=" + cellId;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        bExist = true;
                    }
                }
            }
            return bExist;
        }
        public bool IsPackageExist(int packageId)
        {
            bool bExist = false;
            string SelectCommand = @"SELECT CabFurniturePackageID FROM CabFurniturePackages WHERE CabFurniturePackageID=" + packageId;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        bExist = true;
                    }
                }
            }
            return bExist;
        }

        /// <summary>
        /// привязать упаковки к ячейке. Проставляется дата принятия на склад
        /// </summary>
        /// <param name="cellID"></param>
        /// <param name="packageIds"></param>
        public void BindPackagesToCell(int cellID, int[] packageIds)
        {
            string filter = string.Empty;
            foreach (int item in packageIds)
                filter += item + ",";
            if (filter.Length > 0)
                filter = $@"SELECT CompleteClientID, CompleteMegaOrderID, CompleteStorage, CompleteUserID, CompleteDateTime, 
CabFurniturePackageID, CellID, BindToCellUserID, BindToCellDateTime, AddToStorage, AddToStorageDateTime, AddToStorageUserID,
QualityControlOutUserID, QualityControlOutDateTime, QualityControl FROM CabFurniturePackages 
WHERE CabFurniturePackageID IN ({filter.Substring(0, filter.Length - 1)})";

            using (SqlDataAdapter DA = new SqlDataAdapter(filter, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DateTime dateTime = Security.GetCurrentDate();

                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                if (Convert.ToInt32(DT.Rows[i]["QualityControl"]) == 0) // если упаковка находится на ОТК, то принимаем её с ОТК
                                {
                                    DT.Rows[i]["QualityControl"] = 1;
                                    DT.Rows[i]["QualityControlOutDateTime"] = dateTime;
                                    DT.Rows[i]["QualityControlOutUserID"] = Security.CurrentUserID;
                                }
                                DT.Rows[i]["CellID"] = cellID;
                                DT.Rows[i]["CompleteClientID"] = DBNull.Value;
                                DT.Rows[i]["CompleteMegaOrderID"] = DBNull.Value;
                                DT.Rows[i]["CompleteStorage"] = false;
                                DT.Rows[i]["CompleteUserID"] = DBNull.Value;
                                DT.Rows[i]["CompleteDateTime"] = DBNull.Value;
                                DT.Rows[i]["BindToCellUserID"] = Security.CurrentUserID;
                                DT.Rows[i]["BindToCellDateTime"] = dateTime;
                                DT.Rows[i]["AddToStorage"] = 1;
                                DT.Rows[i]["AddToStorageDateTime"] = dateTime;
                                DT.Rows[i]["AddToStorageUserID"] = Security.CurrentUserID;
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public void QualityControlIn(int[] packageIds)
        {
            string filter = string.Empty;
            foreach (int item in packageIds)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = "SELECT CabFurniturePackageID, QualityControlInUserID, QualityControlInDateTime, QualityControl FROM CabFurniturePackages " +
                    "WHERE CabFurniturePackageID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            using (SqlDataAdapter da = new SqlDataAdapter(filter, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (DataTable dt = new DataTable())
                    {
                        if (da.Fill(dt) <= 0) return;

                        DateTime dateTime = Security.GetCurrentDate();

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            dt.Rows[i]["QualityControlInDateTime"] = dateTime;
                            dt.Rows[i]["QualityControlInUserID"] = Security.CurrentUserID;
                            dt.Rows[i]["QualityControl"] = 0;
                        }
                        da.Update(dt);
                    }
                }
            }
        }
        
        public void QualityControlOut(int[] packageIds)
        {
            string filter = string.Empty;
            foreach (int item in packageIds)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = "SELECT CabFurniturePackageID, QualityControlOutUserID, QualityControlOutDateTime, QualityControl FROM CabFurniturePackages " +
                    "WHERE CabFurniturePackageID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            using (SqlDataAdapter da = new SqlDataAdapter(filter, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(da))
                {
                    using (DataTable dt = new DataTable())
                    {
                        if (da.Fill(dt) <= 0) return;

                        DateTime dateTime = Security.GetCurrentDate();

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            dt.Rows[i]["QualityControlOutDateTime"] = dateTime;
                            dt.Rows[i]["QualityControlOutUserID"] = Security.CurrentUserID;
                            dt.Rows[i]["QualityControl"] = 1;
                        }
                        da.Update(dt);
                    }
                }
            }
        }

        /// <summary>
        /// отвязать упаковки от ячейки
        /// </summary>
        /// <param name="cellID"></param>
        /// <param name="packageIds"></param>
        public void UnbindPackagesToCell(int[] packageIds)
        {
            string filter = string.Empty;
            foreach (int item in packageIds)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = "SELECT CabFurniturePackageID, CellID, BindToCellUserID, BindToCellDateTime, RemoveFromStorage, RemoveFromStorageDateTime, RemoveFromStorageUserID FROM CabFurniturePackages " +
                    "WHERE CabFurniturePackageID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(filter, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DateTime dateTime = Security.GetCurrentDate();

                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                DT.Rows[i]["CellID"] = -1;
                                DT.Rows[i]["RemoveFromStorage"] = 1;
                                DT.Rows[i]["RemoveFromStorageDateTime"] = dateTime;
                                DT.Rows[i]["RemoveFromStorageUserID"] = Security.CurrentUserID;
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
        }

    }

    public class AssemblePackagesManager
    {
        private int iScanedPackages = 0;
        private int iAllPackages = 0;

        private string sRackName = string.Empty;
        private string sCellName = string.Empty;

        DataTable ScanedPackagesDT = null;

        DataTable NotScanedPackageLabelsDT = null;
        DataTable AllScanedPackageLabelsDT = null;
        public BindingSource NotScanedPackageLabelsBS = null;
        public BindingSource AllScanedPackageLabelsBS = null;
        SqlDataAdapter NotScanedPackageLabelsDA;

        DataTable NotScanedPackageDetailsDT = null;
        DataTable AllScanedPackageDetailsDT = null;
        public BindingSource NotScanedPackageDetailsBS = null;
        public BindingSource AllScanedPackageDetailsBS = null;
        SqlDataAdapter NotScanedPackageDetailsDA;

        DataTable ScanedPackageDetailsDT = null;
        public BindingSource ScanedPackageDetailsBS = null;
        SqlDataAdapter ScanedPackageDetailsDA;

        public string ScanedPackages { get => iScanedPackages + "/" + iAllPackages; }
        public string CellName { get => sCellName; set => sCellName = value; }
        public string RackName { get => sRackName; set => sRackName = value; }

        public AssemblePackagesManager()
        {
            ScanedPackagesDT = new DataTable();
            ScanedPackagesDT.Columns.Add(new DataColumn("CabFurniturePackageID", Type.GetType("System.Int32")));

            NotScanedPackageLabelsDT = new DataTable();
            AllScanedPackageLabelsDT = new DataTable();
            NotScanedPackageDetailsDT = new DataTable();
            AllScanedPackageDetailsDT = new DataTable();
            ScanedPackageDetailsDT = new DataTable();

            NotScanedPackageLabelsBS = new BindingSource();
            AllScanedPackageLabelsBS = new BindingSource();
            NotScanedPackageDetailsBS = new BindingSource();
            AllScanedPackageDetailsBS = new BindingSource();
            ScanedPackageDetailsBS = new BindingSource();

            string SelectCommand = @"SELECT TOP 0 CabFurnitureComplementID, TechCatalogOperationsDetailID, TechStoreSubGroupID, TechStoreID, CoverID, PatinaID, InsetColorID, PackNumber, MainOrderID, Notes FROM CabFurnitureComplements";
            NotScanedPackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            NotScanedPackageLabelsDA.Fill(NotScanedPackageLabelsDT);
            NotScanedPackageLabelsDT.Columns.Add(new DataColumn("Index", Type.GetType("System.Int32")));
            NotScanedPackageLabelsDT.Columns.Add(new DataColumn("CabFurniturePackageID", Type.GetType("System.Int32")));
            NotScanedPackageLabelsDT.Columns.Add(new DataColumn("Scan", Type.GetType("System.Boolean")));
            NotScanedPackageLabelsDT.Columns["Scan"].DefaultValue = false;

            AllScanedPackageLabelsDT = NotScanedPackageLabelsDT.Clone();

            SelectCommand = @"SELECT TOP 0 C.PackNumber, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID, Racks.Name AS RackName, Cells.Name AS CellName, CabFurniturePackageDetails.* FROM CabFurniturePackageDetails 
                INNER JOIN CabFurniturePackages AS C ON CabFurniturePackageDetails.CabFurniturePackageID=C.CabFurniturePackageID
                LEFT JOIN Cells ON C.CellID=Cells.CellID
                LEFT JOIN Racks ON Cells.RackID=Racks.RackID";
            ScanedPackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            ScanedPackageDetailsDA.Fill(ScanedPackageDetailsDT);

            SelectCommand = @"SELECT TOP 0 C.PackNumber, C.MainOrderID, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID,
                CabFurnitureComplementDetails.* FROM CabFurnitureComplementDetails 
                INNER JOIN CabFurnitureComplements AS C ON CabFurnitureComplementDetails.CabFurnitureComplementID=C.CabFurnitureComplementID";
            //NotScanedPackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            //NotScanedPackageDetailsDA.Fill(NotScanedPackageDetailsDT);

            NotScanedPackageDetailsDT = GetDataTableByAdapter(SelectCommand);
            NotScanedPackageDetailsDT.Columns.Add(new DataColumn("Scan", Type.GetType("System.Boolean")));
            NotScanedPackageDetailsDT.Columns.Add(new DataColumn("CabFurniturePackageID", Type.GetType("System.Int32")));
            NotScanedPackageDetailsDT.Columns["Scan"].DefaultValue = false;

            AllScanedPackageDetailsDT = NotScanedPackageDetailsDT.Clone();

            NotScanedPackageLabelsBS = new BindingSource()
            {
                DataSource = new DataView(NotScanedPackageLabelsDT, "Scan=0", string.Empty, DataViewRowState.CurrentRows)
            };

            AllScanedPackageLabelsBS.DataSource = AllScanedPackageLabelsDT;
            NotScanedPackageDetailsBS.DataSource = NotScanedPackageDetailsDT;
            AllScanedPackageDetailsBS.DataSource = AllScanedPackageDetailsDT;
            ScanedPackageDetailsBS.DataSource = ScanedPackageDetailsDT;
        }

        public void Clear()
        {
            iScanedPackages = 0;
            iAllPackages = 0;
            ScanedPackagesDT.Clear();
            NotScanedPackageDetailsDT.Clear();
            AllScanedPackageDetailsDT.Clear();
        }

        public DataTable GetDataTableByAdapter(string query)
        {
            DataTable dt = new DataTable();
            using (SqlConnection sqlConn = new SqlConnection(ConnectionStrings.StorageConnectionString))
            using (SqlCommand cmd = new SqlCommand(query, sqlConn))
            {
                DateTime start = DateTime.Now;
                sqlConn.Open();
                new SqlDataAdapter(query, sqlConn).Fill(dt);
                TimeSpan ts = DateTime.Now.Subtract(start);
                System.Diagnostics.Trace.Write("Time Elapsed in Adapter: " + ts.TotalMilliseconds);
            }

            return dt;
        }

        public DataTable GetDataTableFromDataReader(IDataReader dataReader)
        {
            DataTable schemaTable = dataReader.GetSchemaTable();
            DataTable resultTable = new DataTable();

            foreach (DataRow dataRow in schemaTable.Rows)
            {
                DataColumn dataColumn = new DataColumn();
                dataColumn.ColumnName = dataRow["ColumnName"].ToString();
                dataColumn.DataType = Type.GetType(dataRow["DataType"].ToString());
                dataColumn.ReadOnly = (bool)dataRow["IsReadOnly"];
                dataColumn.AutoIncrement = (bool)dataRow["IsAutoIncrement"];
                dataColumn.Unique = (bool)dataRow["IsUnique"];

                resultTable.Columns.Add(dataColumn);
            }

            while (dataReader.Read())
            {
                DataRow dataRow = resultTable.NewRow();
                for (int i = 0; i < resultTable.Columns.Count - 1; i++)
                {
                    dataRow[i] = dataReader[i];
                }
                resultTable.Rows.Add(dataRow);
            }

            return resultTable;
        }

        public bool GetPackagesLabels(int MegaOrderID)
        {
            string SelectCommand = @"SELECT C.PackNumber, C.MainOrderID, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID,
                CabFurnitureComplementDetails.*, 0 as Scan, NULL as CabFurniturePackageID FROM CabFurnitureComplementDetails 
                INNER JOIN CabFurnitureComplements AS C ON CabFurnitureComplementDetails.CabFurnitureComplementID=C.CabFurnitureComplementID
                WHERE MainOrderID IN (SELECT MainOrderID FROM infiniu2_marketingorders.dbo.MainOrders WHERE MegaOrderID = " + MegaOrderID + ")";
            //NotScanedPackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            //NotScanedPackageDetailsDA.Fill(NotScanedPackageDetailsDT);
            NotScanedPackageDetailsDT = GetDataTableByAdapter(SelectCommand);
            //NotScanedPackageDetailsDT.Columns.Add(new DataColumn("Scan", Type.GetType("System.Boolean")));
            //NotScanedPackageDetailsDT.Columns.Add(new DataColumn("CabFurniturePackageID", Type.GetType("System.Int32")));
            //NotScanedPackageDetailsDT.Columns["Scan"].DefaultValue = false;

            //using (SqlConnection sqlConn = new SqlConnection(ConnectionStrings.StorageConnectionString))
            //using (SqlCommand cmd = new SqlCommand(SelectCommand, sqlConn))
            //{
            //    sqlConn.Open();
            //    NotScanedPackageDetailsDT.Load(cmd.ExecuteReader());
            //}

            NotScanedPackageLabelsDT.Clear();
            DataTable dt = NotScanedPackageLabelsDT.Clone();
            SelectCommand = @"SELECT CabFurnitureComplementID, TechCatalogOperationsDetailID, TechStoreSubGroupID, TechStoreID, 
CoverID, PatinaID, InsetColorID, PackNumber, MainOrderID, Notes FROM CabFurnitureComplements 
WHERE MainOrderID IN (SELECT MainOrderID FROM infiniu2_marketingorders.dbo.MainOrders WHERE MegaOrderID =" + MegaOrderID + ")";
            NotScanedPackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            NotScanedPackageLabelsDA.Fill(dt);
            
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //dt.Rows[i]["Scan"] = 0;
                dt.Rows[i]["Index"] = i + 1;
            }

            foreach (DataRow dr in dt.Rows)
                NotScanedPackageLabelsDT.Rows.Add(dr.ItemArray);
            dt.Dispose();
            iAllPackages = NotScanedPackageLabelsDT.Rows.Count;

            return NotScanedPackageLabelsDT.Rows.Count > 0;
        }

        public bool GetPackagesLabels(int ClientID, int MegaOrderID)
        {
            string SelectCommand = @"SELECT C.PackNumber, C.MainOrderID, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID,
                CabFurnitureComplementDetails.*, 0 as Scan, NULL as CabFurniturePackageID FROM CabFurnitureComplementDetails 
                INNER JOIN CabFurnitureComplements AS C ON CabFurnitureComplementDetails.CabFurnitureComplementID=C.CabFurnitureComplementID
                WHERE ClientID=" + ClientID + " AND MegaOrderID=" + MegaOrderID;
            NotScanedPackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            NotScanedPackageDetailsDA.Fill(NotScanedPackageDetailsDT);

            for (int i = 0; i < NotScanedPackageDetailsDT.Rows.Count; i++)
            {
                NotScanedPackageDetailsDT.Rows[i]["Scan"] = 0;
            }

            NotScanedPackageLabelsDT.Clear();
            DataTable dt = NotScanedPackageLabelsDT.Clone();
            SelectCommand = @"SELECT CabFurnitureComplementID, TechCatalogOperationsDetailID, TechStoreSubGroupID, TechStoreID, 
CoverID, PatinaID, InsetColorID, PackNumber, MainOrderID, Notes FROM CabFurnitureComplements WHERE ClientID=" + ClientID + " AND MegaOrderID=" + MegaOrderID;
            NotScanedPackageLabelsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            NotScanedPackageLabelsDA.Fill(dt);
            
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["Scan"] = 0;
                dt.Rows[i]["Index"] = i + 1;
            }

            foreach (DataRow dr in dt.Rows)
                NotScanedPackageLabelsDT.Rows.Add(dr.ItemArray);
            dt.Dispose();
            iAllPackages = NotScanedPackageLabelsDT.Rows.Count;

            return NotScanedPackageLabelsDT.Rows.Count > 0;
        }

        public void FilterNotScanedPackagesDetails(int CabFurnitureComplementID)
        {
            NotScanedPackageDetailsBS.Filter = "CabFurnitureComplementID =" + CabFurnitureComplementID;
            NotScanedPackageDetailsBS.MoveFirst();
        }

        public void FilterAllScanedPackagesDetails(int CabFurnitureComplementID)
        {
            AllScanedPackageDetailsBS.Filter = "CabFurnitureComplementID =" + CabFurnitureComplementID;
            AllScanedPackageDetailsBS.MoveFirst();
        }

        public DataTable ScanedPackagesToExport
        {
            get
            {
                return NotScanedPackageDetailsDT;
            }
        }

        /// <summary>
        /// Проверяет, сканировалась ли упаковка ранее
        /// </summary>
        /// <param name="CabFurniturePackageID">id упаковки</param>
        /// <returns>true-сканировалась, false-еще не сканировалась</returns>
        public bool IsPackageScan(int CabFurniturePackageID)
        {
            DataRow[] rows = ScanedPackagesDT.Select("CabFurniturePackageID=" + CabFurniturePackageID);
            return rows.Count() > 0;
        }

        private void AddToScaned(int CabFurniturePackageID)
        {
            DataRow NewRow = ScanedPackagesDT.NewRow();
            NewRow["CabFurniturePackageID"] = CabFurniturePackageID;
            ScanedPackagesDT.Rows.Add(NewRow);
        }

        public class packInfo
        {
            public int CabFurnitureComplementID;
            public bool Scan;
            public int CabFurniturePackageID;
        }

        /// <summary>
        /// Проверяет, соответствует ли упаковка продукту в заказе
        /// </summary>
        /// <param name="TechCatalogOperationsDetailID"></param>
        /// <param name="TechStoreID"></param>
        /// <param name="CoverID"></param>
        /// <param name="PatinaID"></param>
        /// <param name="InsetColorID"></param>
        /// <returns></returns>
        public packInfo IsPackageMatch(int CabFurniturePackageID, int TechCatalogOperationsDetailID, int TechStoreID, int CoverID, int PatinaID, int InsetColorID)
        {
            packInfo p = new packInfo();
            bool b = false;
            DataRow[] rows = NotScanedPackageLabelsDT.Select("Scan=0 AND TechCatalogOperationsDetailID=" + TechCatalogOperationsDetailID +
                " AND TechStoreID=" + TechStoreID +
                " AND CoverID=" + CoverID +
                " AND PatinaID=" + PatinaID +
                " AND InsetColorID=" + InsetColorID);
            if (rows.Count() > 0)
            {
                int CabFurnitureComplementID = Convert.ToInt32(rows[0]["CabFurnitureComplementID"]);

                //NotScanedPackageLabelsDT.Select("CabFurnitureComplementID=" + CabFurnitureComplementID)[0].SetField("Scan", 1);
                //NotScanedPackageLabelsDT.Select("CabFurnitureComplementID=" + CabFurnitureComplementID)[0].SetField("CabFurniturePackageID", CabFurniturePackageID);

                rows[0]["Scan"] = 1;
                rows[0]["CabFurniturePackageID"] = CabFurniturePackageID;
                NotScanedPackageLabelsDT.AcceptChanges();
                foreach (DataRow dr1 in NotScanedPackageDetailsDT.Select("CabFurnitureComplementID=" + CabFurnitureComplementID))
                {
                    dr1["Scan"] = 1;
                    dr1["CabFurniturePackageID"] = CabFurniturePackageID;
                    AllScanedPackageDetailsDT.Rows.Add(dr1.ItemArray);
                }
                AllScanedPackageLabelsDT.Rows.Add(rows[0].ItemArray);

                //rows.AsEnumerable()
                //    .Where(row => row.Field<Int64>("CabFurnitureComplementID") == CabFurnitureComplementID)
                //    .ToList<DataRow>()
                //    .ForEach(row => { row["CabFurniturePackageID"] = CabFurniturePackageID; row["Scan"] = 1; });

                p.CabFurnitureComplementID = CabFurnitureComplementID;
                p.Scan = true;
                p.CabFurniturePackageID = CabFurniturePackageID;

            }

            //foreach (DataRow dr in NotScanedPackageLabelsDT.Rows)
            //{
            //    if (Convert.ToInt32(dr["TechCatalogOperationsDetailID"]) == TechCatalogOperationsDetailID &&
            //        Convert.ToInt32(dr["TechStoreID"]) == TechStoreID &&
            //        Convert.ToInt32(dr["CoverID"]) == CoverID &&
            //        Convert.ToInt32(dr["PatinaID"]) == PatinaID &&
            //        Convert.ToInt32(dr["InsetColorID"]) == InsetColorID &&
            //        Convert.ToBoolean(dr["Scan"]) == false)
            //    {
            //        int CabFurnitureComplementID = Convert.ToInt32(dr["CabFurnitureComplementID"]);

            //        dr["Scan"] = 1;
            //        dr["CabFurniturePackageID"] = CabFurniturePackageID;
            //        foreach (DataRow dr1 in NotScanedPackageDetailsDT.Select("CabFurnitureComplementID=" + CabFurnitureComplementID))
            //        {
            //            AllScanedPackageDetailsDT.Rows.Add(dr1.ItemArray);
            //        }
            //        AllScanedPackageLabelsDT.Rows.Add(dr.ItemArray);
            //        b = true;
            //        break;
            //    }
            //}
            return p;
        }

        /// <summary>
        /// Сканирует этикетку, проверяет упаковку, добавляет id упаковки в таблицу отсканированных
        /// </summary>
        /// <param name="CabFurniturePackageID"></param>
        /// <returns>false, если упаковка не соответствует</returns>
        public bool ScanPackage(int CabFurniturePackageID)
        {
            bool b = false;

            string SelectCommand = @"SELECT CabFurniturePackageID, CabFurAssignmentID, CabFurAssignmentDetailID, 
CompleteClientID, CompleteMegaOrderID,
TechCatalogOperationsDetailID, PackNumber, ClientName, TechStoreSubGroupID, TechStoreID, CoverID, PatinaID, InsetColorID, PackagesCount FROM CabFurniturePackages WHERE CabFurniturePackageID=" + CabFurniturePackageID;
            using (SqlDataAdapter da = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                using (DataTable dt = new DataTable())
                {
                    if (da.Fill(dt) > 0)
                    {
                        int TechCatalogOperationsDetailID = Convert.ToInt32(dt.Rows[0]["TechCatalogOperationsDetailID"]);
                        int TechStoreID = Convert.ToInt32(dt.Rows[0]["TechStoreID"]);
                        int CoverID = Convert.ToInt32(dt.Rows[0]["CoverID"]);
                        int PatinaID = Convert.ToInt32(dt.Rows[0]["PatinaID"]);
                        int InsetColorID = Convert.ToInt32(dt.Rows[0]["InsetColorID"]);

                        packInfo p = IsPackageMatch(CabFurniturePackageID, TechCatalogOperationsDetailID, TechStoreID, CoverID, PatinaID, InsetColorID);

                        if (p.Scan)
                        {
                            iScanedPackages++;
                            b = true;
                            AddToScaned(CabFurniturePackageID);
                        }
                    }
                }
            }

            ScanedPackageDetailsDT.Clear();
            SelectCommand = @"SELECT C.PackNumber, C.TechStoreSubGroupID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID, Racks.Name AS RackName, Cells.Name AS CellName, CabFurniturePackageDetails.* FROM CabFurniturePackageDetails 
                INNER JOIN CabFurniturePackages AS C ON CabFurniturePackageDetails.CabFurniturePackageID=C.CabFurniturePackageID
                LEFT JOIN Cells ON C.CellID=Cells.CellID
                LEFT JOIN Racks ON Cells.RackID=Racks.RackID
                WHERE CabFurniturePackageDetails.CabFurniturePackageID=" + CabFurniturePackageID;
            ScanedPackageDetailsDA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString);
            if (ScanedPackageDetailsDA.Fill(ScanedPackageDetailsDT) > 0)
            {
                sRackName = ScanedPackageDetailsDT.Rows[0]["RackName"].ToString();
                sCellName = ScanedPackageDetailsDT.Rows[0]["CellName"].ToString();
            }

            return b;
        }

        public bool IsPackageExist(int CabFurniturePackageID)
        {
            bool b = false;
            string SelectCommand = @"SELECT CabFurniturePackageID FROM CabFurniturePackages WHERE CabFurniturePackageID=" + CabFurniturePackageID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        b = true;
                    }
                }
            }
            return b;
        }

        public void IsPackageCompleteToClient(int CompleteMegaOrderID)
        {
            string SelectCommand = @"SELECT * FROM CabFurniturePackages WHERE CompleteMegaOrderID=" + CompleteMegaOrderID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                using (DataTable dt = new DataTable())
                {
                    if (DA.Fill(dt) > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            int CabFurniturePackageID = Convert.ToInt32(dt.Rows[i]["CabFurniturePackageID"]);
                            int TechCatalogOperationsDetailID = Convert.ToInt32(dt.Rows[i]["TechCatalogOperationsDetailID"]);
                            int TechStoreID = Convert.ToInt32(dt.Rows[i]["TechStoreID"]);
                            int CoverID = Convert.ToInt32(dt.Rows[i]["CoverID"]);
                            int PatinaID = Convert.ToInt32(dt.Rows[i]["PatinaID"]);
                            int InsetColorID = Convert.ToInt32(dt.Rows[i]["InsetColorID"]);

                            packInfo p = IsPackageMatch(CabFurniturePackageID, TechCatalogOperationsDetailID, TechStoreID, CoverID, PatinaID, InsetColorID);
                            if (p.Scan)
                            {
                                iScanedPackages++;
                                AddToScaned(CabFurniturePackageID);
                                //NotScanedPackageLabelsDT.Select("CabFurnitureComplementID=" + p.CabFurnitureComplementID)[0].SetField("Scan", p.Scan);
                                //NotScanedPackageLabelsDT.Select("CabFurnitureComplementID=" + p.CabFurnitureComplementID)[0].SetField("CabFurniturePackageID", p.CabFurniturePackageID);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// упаковка собрана и готова к отгрузке
        /// </summary>
        public void AssemblePackages(int CompleteClientID, int CompleteMegaOrderID)
        {
            string filter = string.Empty;
            foreach (DataRow item in ScanedPackagesDT.Rows)
                filter += item["CabFurniturePackageID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "SELECT CellID, CabFurniturePackageID, CompleteStorage, CompleteClientID, CompleteMegaOrderID, CompleteUserID, CompleteDateTime FROM CabFurniturePackages " +
                    "WHERE CabFurniturePackageID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(filter, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DateTime dateTime = Security.GetCurrentDate();

                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                DT.Rows[i]["CellID"] = -1;
                                DT.Rows[i]["CompleteClientID"] = CompleteClientID;
                                DT.Rows[i]["CompleteMegaOrderID"] = CompleteMegaOrderID;
                                DT.Rows[i]["CompleteStorage"] = 1;
                                DT.Rows[i]["CompleteDateTime"] = dateTime;
                                DT.Rows[i]["CompleteUserID"] = Security.CurrentUserID;
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public void FF()
        {
            string filter1 = "SELECT CabFurnitureComplementID, MegaOrderID, MainOrderID FROM CabFurnitureComplements";
            string filter2 = "SELECT distinct MainOrderID, MegaOrderID FROM infiniu2_marketingorders.dbo.NewMainOrders";

            DataTable dt = new DataTable();

            using (SqlDataAdapter DA = new SqlDataAdapter(filter2, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(dt);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(filter1, ConnectionStrings.StorageConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                int MainOrderID = Convert.ToInt32(DT.Rows[i]["MainOrderID"]);
                                int MegaOrderID = -1;
                                DataRow[] rows = dt.Select("MainOrderID=" + MainOrderID);
                                if (rows.Any())
                                    MegaOrderID = Convert.ToInt32(rows[0]["MegaOrderID"]);
                                DT.Rows[i]["MegaOrderID"] = MegaOrderID;
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
        }
    }



    public class CabFurAssemble
    {
        DataTable AllMegaOrdersDecorWeightDT;
        DataTable AllMainOrdersDecorWeightDT;
        DataTable MegaOrdersDT;
        DataTable AssembleDatesDT;
        DataTable AllMainOrdersDT;
        DataTable MainOrdersDT;

        DataTable CabFurOrdersDataTable;
        DataTable AllCabFurniturePackages;
        DataTable CabFurniturePackages;

        BindingSource MegaOrdersBS;
        BindingSource AssembleDatesBS;
        BindingSource MainOrdersBS;

        public CabFurAssemble()
        {

        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
        }

        private void Create()
        {
            CabFurOrdersDataTable = new DataTable();
            AllCabFurniturePackages = new DataTable();
            AllMainOrdersDT = new DataTable();

            CabFurniturePackages = new DataTable();
            CabFurniturePackages.Columns.Add(new DataColumn(("CabFurniturePackageID"), System.Type.GetType("System.Int32")));
            CabFurniturePackages.Columns.Add(new DataColumn(("CellID"), System.Type.GetType("System.Int32")));
            CabFurniturePackages.Columns.Add(new DataColumn(("CellName"), System.Type.GetType("System.String")));

            AllMegaOrdersDecorWeightDT = new DataTable();
            AllMainOrdersDecorWeightDT = new DataTable();
            MainOrdersDT = new DataTable();
            MegaOrdersDT = new DataTable();
            MegaOrdersDT.Columns.Add(new DataColumn(("PackagesCount"), System.Type.GetType("System.String")));
            MegaOrdersDT.Columns.Add(new DataColumn(("Weight"), System.Type.GetType("System.Decimal")));
            MegaOrdersDT.Columns.Add(new DataColumn(("Status"), System.Type.GetType("System.String")));
            AssembleDatesDT = new DataTable();
            AssembleDatesDT.Columns.Add(new DataColumn(("WeekNumber"), System.Type.GetType("System.String")));

            MegaOrdersBS = new BindingSource();
            AssembleDatesBS = new BindingSource();
            MainOrdersBS = new BindingSource();
        }

        private void Fill()
        {
            string SelectCommand = @"SELECT TOP 0 MegaOrders.MegaOrderID, MegaOrders.ClientID, MegaOrders.OrderDate, MegaOrders.OrderNumber, Clients.ClientName FROM MegaOrders
                LEFT JOIN infiniu2_marketingreference.dbo.Clients AS Clients ON MegaOrders.ClientID = Clients.ClientID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MegaOrdersDT);
            }
            SelectCommand = "SELECT TOP 0 PrepareDateTime, DateName FROM CabFurDispatch ORDER BY PrepareDateTime";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(AssembleDatesDT);
            }
            SelectCommand = "SELECT CabFurniturePackages.*, Cells.Name FROM CabFurniturePackages" +
                " LEFT JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID" +
                " WHERE CabFurniturePackages.CompleteStorage=1 ORDER BY AddToStorageDateTime";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                DA.Fill(AllCabFurniturePackages);
            }

            SelectCommand = "SELECT TOP 0 MainOrders.MegaOrderID, MegaOrders.ClientID, MegaOrders.OrderNumber, MainOrders.FactoryID, MainOrders.MainOrderID FROM MainOrders INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MainOrdersDT);
            }
            MainOrdersDT.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            MainOrdersDT.Columns.Add(new DataColumn("AllPackCount", Type.GetType("System.Int32")));
            MainOrdersDT.Columns.Add(new DataColumn("PackPercentage", Type.GetType("System.Decimal")));
        }

        public void GetMainOrdersSquareAndWeight()
        {
            string filter = string.Empty;

            foreach (int item in Security.CabFurIds)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = "WHERE ProductID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            string SelectCommand = @"SELECT MainOrders.MegaOrderID, DecorOrders.MainOrderID, (DecorOrders.Weight*PackageDetails.Count/DecorOrders.Count) AS Weight
                    FROM PackageDetails INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM DecorOrders " + filter + @" )
                    INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1  
                    INNER JOIN  MainOrders ON Packages.MainOrderID = MainOrders.MainOrderID 
                    AND MainOrders.MegaOrderID NOT IN (SELECT MegaOrderID FROM CabFurDispatch)";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                AllMainOrdersDecorWeightDT.Clear();
                DA.Fill(AllMainOrdersDecorWeightDT);
            }

            SelectCommand = @"SELECT MainOrders.MegaOrderID, (DecorOrders.Weight*PackageDetails.Count/DecorOrders.Count) AS Weight
                    FROM PackageDetails INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM DecorOrders " + filter + @" )
                    INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1  
                    INNER JOIN  MainOrders ON Packages.MainOrderID = MainOrders.MainOrderID 
                    AND MainOrders.MegaOrderID NOT IN (SELECT MegaOrderID FROM CabFurDispatch)";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                AllMegaOrdersDecorWeightDT.Clear();
                DA.Fill(AllMegaOrdersDecorWeightDT);
            }
        }

        public void GetMainOrdersSquareAndWeight(DateTime PrepareDateTime)
        {
            string SelectCommand = @"SELECT MainOrders.MegaOrderID, DecorOrders.MainOrderID, (DecorOrders.Weight*PackageDetails.Count/DecorOrders.Count) AS Weight
                    FROM PackageDetails INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID
                    INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1 
                    INNER JOIN  MainOrders ON Packages.MainOrderID = MainOrders.MainOrderID 
                    AND MainOrders.MegaOrderID IN (SELECT MegaOrderID FROM CabFurDispatch WHERE CAST(PrepareDateTime AS Date) = 
                    '" + PrepareDateTime.ToString("yyyy-MM-dd") + "')";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                AllMainOrdersDecorWeightDT.Clear();
                DA.Fill(AllMainOrdersDecorWeightDT);
            }

            SelectCommand = @"SELECT MainOrders.MegaOrderID, (DecorOrders.Weight*PackageDetails.Count/DecorOrders.Count) AS Weight
                    FROM PackageDetails INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID
                    INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1 
                    INNER JOIN  MainOrders ON Packages.MainOrderID = MainOrders.MainOrderID 
                    AND MainOrders.MegaOrderID IN (SELECT MegaOrderID FROM CabFurDispatch WHERE CAST(PrepareDateTime AS Date) = 
                    '" + PrepareDateTime.ToString("yyyy-MM-dd") + "')";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                AllMegaOrdersDecorWeightDT.Clear();
                DA.Fill(AllMegaOrdersDecorWeightDT);
            }
        }

        private decimal GetWeight(int MegaOrderID, int MainOrderID)
        {
            decimal Weight = 0;
            DataRow[] rows = AllMainOrdersDecorWeightDT.Select("MegaOrderID=" + MegaOrderID + " AND MainOrderID=" + MainOrderID);
            //if (rows.Count() > 0 && rows[0]["Weight"] != DBNull.Value)
            //{
            //    Weight += Convert.ToDecimal(rows[0]["Weight"]);
            //}
            foreach (DataRow item in rows)
            {
                Weight += Convert.ToDecimal(item["Weight"]);
            }
            Weight = Decimal.Round(Weight, 2, MidpointRounding.AwayFromZero);

            return Weight;
        }

        private decimal GetWeight(int MegaOrderID)
        {
            decimal Weight = 0;
            DataRow[] rows = rows = AllMegaOrdersDecorWeightDT.Select("MegaOrderID=" + MegaOrderID);
            foreach (DataRow item in rows)
            {
                Weight += Convert.ToDecimal(item["Weight"]);
            }
            //if (rows.Count() > 0 && rows[0]["Weight"] != DBNull.Value)
            //{
            //    Weight += Convert.ToDecimal(rows[0]["Weight"]);
            //}
            Weight = Decimal.Round(Weight, 2, MidpointRounding.AwayFromZero);

            return Weight;
        }

        public void CreateCabFur()
        {
            string filter = string.Empty;

            foreach (int item in Security.CabFurIds)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = "WHERE ProductID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            DataTable table = new DataTable();
            string SelectCommand = "SELECT TOP 0 * FROM CabFurDispatch";
            using (SqlDataAdapter sda = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(sda))
                {
                    sda.Fill(table);
                    SelectCommand = @"SELECT MegaOrderID, ClientID, ProfilDispatchDate, TPSDispatchDate FROM MegaOrders
                WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM DecorOrders " + filter + " ))";
                    using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    {
                        using (DataTable DT = new DataTable())
                        {
                            DA.Fill(DT);
                            DateTime dateTime = Security.GetCurrentDate();

                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                if (DT.Rows[i]["ProfilDispatchDate"] == DBNull.Value && DT.Rows[i]["TPSDispatchDate"] == DBNull.Value)
                                    continue;

                                DataRow NewRow = table.NewRow();
                                NewRow["CreationDateTime"] = dateTime;
                                if (DT.Rows[i]["ProfilDispatchDate"] != DBNull.Value)
                                    NewRow["PrepareDateTime"] = Convert.ToDateTime(DT.Rows[i]["ProfilDispatchDate"]);
                                else
                                    NewRow["PrepareDateTime"] = Convert.ToDateTime(DT.Rows[i]["TPSDispatchDate"]);
                                NewRow["MegaOrderID"] = Convert.ToInt32(DT.Rows[i]["MegaOrderID"]);
                                NewRow["ClientID"] = Convert.ToInt32(DT.Rows[i]["ClientID"]);
                                table.Rows.Add(NewRow);
                            }
                        }
                    }
                    sda.Update(table);
                }
            }
        }

        public bool FilterCabFurOrders(int[] MegaOrders)
        {
            CabFurniturePackages.Clear();
            string SelectCommand = @"SELECT C.PackNumber, C.TechStoreSubGroupID, C.TechCatalogOperationsDetailID, C.TechStoreID AS CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID, C.MainOrderID, CabFurnitureComplementDetails.* FROM CabFurnitureComplementDetails 
                INNER JOIN CabFurnitureComplements AS C ON CabFurnitureComplementDetails.CabFurnitureComplementID=C.CabFurnitureComplementID
                WHERE MainOrderID IN (SELECT MainOrderID FROM infiniu2_marketingorders.dbo.MainOrders WHERE MegaOrderID IN (" + string.Join(",", MegaOrders) + ")) ORDER BY C.MainOrderID, CTechStoreID, C.CoverID, C.PatinaID, C.InsetColorID, C.CabFurnitureComplementID, C.PackNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                CabFurOrdersDataTable.Clear();
                DA.Fill(CabFurOrdersDataTable);
            }
            FilterMainOrders(MegaOrders);
            return CabFurOrdersDataTable.Rows.Count > 0;
        }
        
        public bool FilterMainOrders(int[] MegaOrders)
        {
            string SelectCommand = @"SELECT MainOrderID, MegaOrderID FROM MainOrders WHERE MegaOrderID IN (" + string.Join(",", MegaOrders) + ") ORDER by MainOrderID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                AllMainOrdersDT.Clear();
                DA.Fill(AllMainOrdersDT);
            }
            return AllMainOrdersDT.Rows.Count > 0;
        }

        private bool IsPackageMatch(int ClientID, int TechCatalogOperationsDetailID, int TechStoreID, int CoverID, int PatinaID, int InsetColorID)
        {
            bool b = false;

            DataRow[] dataRows = AllCabFurniturePackages.Select("CompleteStorage=1 AND CompleteClientID=" + ClientID +" AND TechCatalogOperationsDetailID=" + TechCatalogOperationsDetailID +
                " AND TechStoreID=" + TechStoreID +
                " AND CoverID=" + CoverID +
                " AND PatinaID=" + PatinaID +
                " AND InsetColorID=" + InsetColorID);
            if (dataRows.Count() > 0)
            {
                for (int i = 0; i < dataRows.Count(); i++)
                {
                    int CabFurniturePackageID = Convert.ToInt32(dataRows[i]["CabFurniturePackageID"]);

                    DataRow[] rows = CabFurniturePackages.Select("CabFurniturePackageID = " + CabFurniturePackageID);
                    if (rows.Count() > 0)
                    {
                        continue;
                    }

                    DataRow NewRow = CabFurniturePackages.NewRow();
                    NewRow["CabFurniturePackageID"] = dataRows[i]["CabFurniturePackageID"];
                    NewRow["CellID"] = dataRows[i]["CellID"];
                    NewRow["CellName"] = dataRows[i]["Name"];
                    CabFurniturePackages.Rows.Add(NewRow);

                    b = true;
                    break;
                }
            }

            return b;
        }

        private void FillPackagesInfo()
        {
            int PackedCount = 0;
            int AllCount = 0;
            string Status = string.Empty;

            for (int i = 0; i < MegaOrdersDT.Rows.Count; i++)
            {
                CabFurniturePackages.Clear();
                PackedCount = 0;
                AllCount = 0;
                int MegaOrderID = Convert.ToInt32(MegaOrdersDT.Rows[i]["MegaOrderID"]);
                Tuple<int, int> tuple = GetPackagesInfo(Convert.ToInt32(MegaOrdersDT.Rows[i]["ClientID"]), MegaOrderID);
                PackedCount = tuple.Item1;
                AllCount = tuple.Item2;

                if (AllCount > 0)
                {
                    Status = "Частично скомплектован";
                    if (PackedCount == 0)
                        Status = "Не скомплектован";
                    if (AllCount == PackedCount)
                        Status = "Скомплектован";
                }

                MegaOrdersDT.Rows[i]["Status"] = Status;
                MegaOrdersDT.Rows[i]["PackagesCount"] = PackedCount + " / " + AllCount;
                MegaOrdersDT.Rows[i]["Weight"] = GetWeight(Convert.ToInt32(MegaOrdersDT.Rows[i]["MegaOrderID"]));
            }
        }

        public void FillPercColumns(int ClientID, int MegaOrderID)
        {
            for (int i = 0; i < MainOrdersDT.Rows.Count; i++)
            {
                int MainOrderID = Convert.ToInt32(MainOrdersDT.Rows[i]["MainOrderID"]);
                int PackedCount = 0;
                int AllCount = 0;

                DataTable dt = new DataTable();
                using (DataView DV = new DataView(CabFurOrdersDataTable, "MainOrderID = " + MainOrderID, string.Empty, DataViewRowState.CurrentRows))
                {
                    dt = DV.ToTable(true, new string[] { "CabFurnitureComplementID" });
                }
                if (dt.Rows.Count > 0)
                {
                    int CabFurnitureComplementID = Convert.ToInt32(dt.Rows[0]["CabFurnitureComplementID"]);
                    DataRow[] r = CabFurOrdersDataTable.Select("CabFurnitureComplementID=" + CabFurnitureComplementID);
                    if (IsPackageMatch(ClientID, Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]),
                        Convert.ToInt32(r[0]["CTechStoreID"]),
                        Convert.ToInt32(r[0]["CoverID"]),
                        Convert.ToInt32(r[0]["PatinaID"]),
                        Convert.ToInt32(r[0]["InsetColorID"])))
                        PackedCount++;

                    for (int j = 1; j < dt.Rows.Count; j++)
                    {
                        if (Convert.ToInt32(dt.Rows[j]["CabFurnitureComplementID"]) != CabFurnitureComplementID)
                        {
                            CabFurnitureComplementID = Convert.ToInt32(dt.Rows[j]["CabFurnitureComplementID"]);
                            r = CabFurOrdersDataTable.Select("CabFurnitureComplementID=" + CabFurnitureComplementID);

                            int TechCatalogOperationsDetailID = Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]);
                            int TechStoreID = Convert.ToInt32(r[0]["CTechStoreID"]);
                            int CoverID = Convert.ToInt32(r[0]["CoverID"]);
                            int PatinaID = Convert.ToInt32(r[0]["PatinaID"]);
                            int InsetColorID = Convert.ToInt32(r[0]["InsetColorID"]);

                            bool b = IsPackageMatch(ClientID, TechCatalogOperationsDetailID, TechStoreID, CoverID, PatinaID, InsetColorID);
                            if (b)
                                PackedCount++;
                        }
                    }
                }
                AllCount = dt.Rows.Count;

                decimal PackProgressVal = 0;

                if (AllCount > 0)
                    PackProgressVal = Convert.ToDecimal(Convert.ToDecimal(PackedCount) / Convert.ToDecimal(AllCount));

                decimal d1 = PackProgressVal * 100;
                decimal PackPercentage = Decimal.Round(d1, 1, MidpointRounding.AwayFromZero);

                MainOrdersDT.Rows[i]["Weight"] = GetWeight(MegaOrderID, MainOrderID);
                MainOrdersDT.Rows[i]["AllPackCount"] = AllCount;
                MainOrdersDT.Rows[i]["PackPercentage"] = PackPercentage;
            }
        }

        private void Binding()
        {
            MegaOrdersBS.DataSource = MegaOrdersDT;
            AssembleDatesBS.DataSource = AssembleDatesDT;
            MainOrdersBS.DataSource = MainOrdersDT;
        }

        public object CurrentDate
        {
            get
            {
                if (AssembleDatesBS.Count == 0 || ((DataRowView)AssembleDatesBS.Current).Row["PrepareDateTime"] == DBNull.Value)
                    return DBNull.Value;
                else
                    return ((DataRowView)AssembleDatesBS.Current).Row["PrepareDateTime"];
            }
        }

        public BindingSource MegaOrdersList
        {
            get { return MegaOrdersBS; }
        }

        public BindingSource AssembleDatesList
        {
            get { return AssembleDatesBS; }
        }

        public BindingSource MainOrdersList
        {
            get { return MainOrdersBS; }
        }

        public bool HasPackages(int MegaOrderID)
        {
            string SelectCommand = @"SELECT PackageID FROM Packages WHERE MainOrderID = (SELECT TOP 1 MainOrderID FROM MainOrders WHERE MegaOrderID=" + MegaOrderID + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    return (DA.Fill(DT) > 0);
                }
            }
        }

        private Tuple<int, int> GetPackagesInfo(int ClientID, int MegaOrderID)
        {
            int PackedCount = 0;
            int AllCount = 0;
            DataRow[] mRows = AllMainOrdersDT.Select("MegaOrderID=" + MegaOrderID);

            for (int i = 0; i < mRows.Length; i++)
            {
                int MainOrderID = Convert.ToInt32(mRows[i]["MainOrderID"]);

                //DataTable dt = new DataTable();
                //using (DataView DV = new DataView(CabFurOrdersDataTable, "MainOrderID = " + MainOrderID, string.Empty, DataViewRowState.CurrentRows))
                //{
                //    dt = DV.ToTable(true, new string[] { "CabFurnitureComplementID" });
                //}

                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                sw.Start();
                //CabFurOrdersDataTable.DefaultView.RowFilter = "MainOrderID = " + MainOrderID;
                //dt = CabFurOrdersDataTable.DefaultView.ToTable(true, "CabFurnitureComplementID");

                sw.Stop();
                double G1 = sw.Elapsed.TotalMilliseconds;
                sw.Restart();
                List<Int64> distCabFurnitureComplementIDs = CabFurOrdersDataTable.AsEnumerable()
         .Where(r => r.Field<Int64>("MainOrderID") == MainOrderID) // filter rows by date
         .Select(r => r.Field<Int64>("CabFurnitureComplementID")) // select only wellbore string value
         .Distinct() // take only unique items
         .ToList();

                sw.Stop();
                double G2 = sw.Elapsed.TotalMilliseconds;

                //if (dt.Rows.Count > 0)
                //{
                //    int CabFurnitureComplementID = Convert.ToInt32(dt.Rows[0]["CabFurnitureComplementID"]);
                //    DataRow[] r = CabFurOrdersDataTable.Select("CabFurnitureComplementID=" + CabFurnitureComplementID);
                //    int TechCatalogOperationsDetailID = Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]);
                //    if (IsPackageMatch(ClientID, Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]),
                //        Convert.ToInt32(r[0]["CTechStoreID"]),
                //        Convert.ToInt32(r[0]["CoverID"]),
                //        Convert.ToInt32(r[0]["PatinaID"]),
                //        Convert.ToInt32(r[0]["InsetColorID"])))
                //        PackedCount++;

                //    for (int j = 1; j < dt.Rows.Count; j++)
                //    {
                //        if (Convert.ToInt32(dt.Rows[j]["CabFurnitureComplementID"]) != CabFurnitureComplementID)
                //        {
                //            CabFurnitureComplementID = Convert.ToInt32(dt.Rows[j]["CabFurnitureComplementID"]);
                //            r = CabFurOrdersDataTable.Select("CabFurnitureComplementID=" + CabFurnitureComplementID);

                //            TechCatalogOperationsDetailID = Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]);
                //            int TechStoreID = Convert.ToInt32(r[0]["CTechStoreID"]);
                //            int CoverID = Convert.ToInt32(r[0]["CoverID"]);
                //            int PatinaID = Convert.ToInt32(r[0]["PatinaID"]);
                //            int InsetColorID = Convert.ToInt32(r[0]["InsetColorID"]);

                //            bool b = IsPackageMatch(ClientID, TechCatalogOperationsDetailID, TechStoreID, CoverID, PatinaID, InsetColorID);
                //            if (b)
                //                PackedCount++;
                //        }
                //    }
                //}
                //AllCount += dt.Rows.Count;

                if (distCabFurnitureComplementIDs.Count > 0)
                {
                    int CabFurnitureComplementID = Convert.ToInt32(distCabFurnitureComplementIDs[0]);
                    DataRow[] r = CabFurOrdersDataTable.Select("CabFurnitureComplementID=" + CabFurnitureComplementID);
                    int TechCatalogOperationsDetailID = Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]);
                    if (IsPackageMatch(ClientID, Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]),
                        Convert.ToInt32(r[0]["CTechStoreID"]),
                        Convert.ToInt32(r[0]["CoverID"]),
                        Convert.ToInt32(r[0]["PatinaID"]),
                        Convert.ToInt32(r[0]["InsetColorID"])))
                        PackedCount++;

                    for (int j = 1; j < distCabFurnitureComplementIDs.Count; j++)
                    {
                        if (distCabFurnitureComplementIDs[j] != CabFurnitureComplementID)
                        {
                            CabFurnitureComplementID = Convert.ToInt32(distCabFurnitureComplementIDs[j]);
                            r = CabFurOrdersDataTable.Select("CabFurnitureComplementID=" + CabFurnitureComplementID);

                            TechCatalogOperationsDetailID = Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]);
                            int TechStoreID = Convert.ToInt32(r[0]["CTechStoreID"]);
                            int CoverID = Convert.ToInt32(r[0]["CoverID"]);
                            int PatinaID = Convert.ToInt32(r[0]["PatinaID"]);
                            int InsetColorID = Convert.ToInt32(r[0]["InsetColorID"]);

                            bool b = IsPackageMatch(ClientID, TechCatalogOperationsDetailID, TechStoreID, CoverID, PatinaID, InsetColorID);
                            if (b)
                                PackedCount++;
                        }
                    }
                }
                AllCount += distCabFurnitureComplementIDs.Count;

                decimal PackProgressVal = 0;

                if (AllCount > 0)
                    PackProgressVal = Convert.ToDecimal(Convert.ToDecimal(PackedCount) / Convert.ToDecimal(AllCount));

                decimal d1 = PackProgressVal * 100;
                decimal PackPercentage = Decimal.Round(d1, 1, MidpointRounding.AwayFromZero);
            }
            Tuple<int, int> tuple = new Tuple<int, int>(PackedCount, AllCount);
            return tuple;
        }

        public bool IsCabFurAssembled(int ClientID, int MegaOrderID)
        {
            int PackedCount = 0;
            int AllCount = 0;

            string SelectCommand = @"SELECT DISTINCT MainOrderID FROM MainOrders WHERE MegaOrderID = " + MegaOrderID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        for (int i = 0; i < DT.Rows.Count; i++)
                        {

                            int MainOrderID = Convert.ToInt32(DT.Rows[i]["MainOrderID"]);

                            DataTable dt = new DataTable();
                            using (DataView DV = new DataView(CabFurOrdersDataTable, "MainOrderID = " + MainOrderID, string.Empty, DataViewRowState.CurrentRows))
                            {
                                dt = DV.ToTable(true, new string[] { "CabFurnitureComplementID" });
                            }
                            if (dt.Rows.Count > 0)
                            {
                                int CabFurnitureComplementID = Convert.ToInt32(dt.Rows[0]["CabFurnitureComplementID"]);
                                DataRow[] r = CabFurOrdersDataTable.Select("CabFurnitureComplementID=" + CabFurnitureComplementID);
                                int TechCatalogOperationsDetailID = Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]);
                                if (IsPackageMatch(ClientID, Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]),
                                    Convert.ToInt32(r[0]["CTechStoreID"]),
                                    Convert.ToInt32(r[0]["CoverID"]),
                                    Convert.ToInt32(r[0]["PatinaID"]),
                                    Convert.ToInt32(r[0]["InsetColorID"])))
                                    PackedCount++;

                                for (int j = 1; j < dt.Rows.Count; j++)
                                {
                                    if (Convert.ToInt32(dt.Rows[j]["CabFurnitureComplementID"]) != CabFurnitureComplementID)
                                    {
                                        CabFurnitureComplementID = Convert.ToInt32(dt.Rows[j]["CabFurnitureComplementID"]);
                                        r = CabFurOrdersDataTable.Select("CabFurnitureComplementID=" + CabFurnitureComplementID);

                                        TechCatalogOperationsDetailID = Convert.ToInt32(r[0]["TechCatalogOperationsDetailID"]);
                                        int TechStoreID = Convert.ToInt32(r[0]["CTechStoreID"]);
                                        int CoverID = Convert.ToInt32(r[0]["CoverID"]);
                                        int PatinaID = Convert.ToInt32(r[0]["PatinaID"]);
                                        int InsetColorID = Convert.ToInt32(r[0]["InsetColorID"]);

                                        bool b = IsPackageMatch(ClientID, TechCatalogOperationsDetailID, TechStoreID, CoverID, PatinaID, InsetColorID);
                                        if (b)
                                            PackedCount++;
                                    }
                                }
                            }
                            AllCount += dt.Rows.Count;
                        }
                    }
                }
            }
            bool bAssembled = PackedCount > 0 && AllCount > 0 && PackedCount == AllCount;
            return bAssembled;
        }

        private int GetWeekNumber(DateTime dtPassed)
        {
            System.Globalization.CultureInfo ciCurr = System.Globalization.CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dtPassed, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            return weekNum;
        }

        private void FillWeekNumber()
        {
            for (int i = 0; i < AssembleDatesDT.Rows.Count; i++)
            {
                if (AssembleDatesDT.Rows[i]["PrepareDateTime"] != DBNull.Value)
                    AssembleDatesDT.Rows[i]["WeekNumber"] = GetWeekNumber(Convert.ToDateTime(AssembleDatesDT.Rows[i]["PrepareDateTime"])) + " к.н.";
            }
        }

        public void ClearMegaOrders()
        {
            CabFurniturePackages.Clear();
            MegaOrdersDT.Clear();
        }

        public void ClearAssembleDates()
        {
            CabFurniturePackages.Clear();
            AssembleDatesDT.Clear();
        }

        public void ClearMainOrders()
        {
            CabFurniturePackages.Clear();
            MainOrdersDT.Clear();
        }

        public void ChangeAssembleDate(int MegaOrderID, int ClientID, object PrepareDateTime)
        {
            string SelectCommand = @"SELECT CabFurDispatchID, MegaOrderID, ClientID, CreationDateTime, PrepareDateTime, DateName FROM CabFurDispatch WHERE MegaOrderID=" + MegaOrderID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            if (PrepareDateTime != null)
                            {
                                DT.Rows[0]["PrepareDateTime"] = Convert.ToDateTime(PrepareDateTime);
                            }
                            DA.Update(DT);
                        }
                        else
                        {

                            DataRow NewRow = DT.NewRow();
                            NewRow["PrepareDateTime"] = Convert.ToDateTime(PrepareDateTime);
                            NewRow["CreationDateTime"] = Security.GetCurrentDate();
                            NewRow["MegaOrderID"] = MegaOrderID;
                            NewRow["ClientID"] = ClientID;
                            DT.Rows.Add(NewRow);
                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public void UpdateAllCabFurniturePackages()
        {
            string SelectCommand = "SELECT CabFurniturePackages.*, Cells.Name FROM CabFurniturePackages" +
                " LEFT JOIN Cells ON CabFurniturePackages.CellID=Cells.CellID" +
                " WHERE CabFurniturePackages.CompleteStorage=1 ORDER BY AddToStorageDateTime";
            AllCabFurniturePackages.Clear();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.StorageConnectionString))
            {
                DA.Fill(AllCabFurniturePackages);
            }
        }

        public void UpdateAssembleDates(DateTime Date)
        {
            string SelectCommand = "SELECT DISTINCT PrepareDateTime, DateName FROM CabFurDispatch" +
                " WHERE DATEPART(month, PrepareDateTime) = DATEPART(month, '" + Date.ToString("yyyy-MM-dd") +
                "') AND DATEPART(year, PrepareDateTime) = DATEPART(year, '" + Date.ToString("yyyy-MM-dd") + "')" +
                " ORDER BY PrepareDateTime DESC";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(AssembleDatesDT);
            }

            DataTable dt = AssembleDatesDT.Clone();
            SelectCommand = "SELECT PrepareDateTime, DateName FROM CabFurDispatch" +
                " WHERE CabFurDispatchID = 0";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(dt);
            }

            foreach (DataRow dr in dt.Rows)
                AssembleDatesDT.Rows.Add(dr.ItemArray);
            dt.Dispose();

            FillWeekNumber();
        }

        public bool IsCabFur(int MegaOrderID)
        {
            string filter = string.Empty;

            foreach (int item in Security.CabFurIds)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = "WHERE ProductID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            string SelectCommand = @"SELECT MainOrderID FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM DecorOrders " + filter + " ) AND MegaOrderID=" + MegaOrderID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return true;
                }
            }
            return false;
        }

        public void FilterAssembleByID(int CabFurDispatchID)
        {
            DataTable dt = MegaOrdersDT.Clone();

            string filter = string.Empty;

            foreach (int item in Security.CabFurIds)
                filter += item.ToString() + ",";
            if (filter.Length > 0)
                filter = "WHERE ProductID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            string SelectCommand = @"SELECT MegaOrders.MegaOrderID, MegaOrders.ClientID, MegaOrders.OrderDate, MegaOrders.OrderNumber, Clients.ClientName FROM MegaOrders
                LEFT JOIN infiniu2_marketingreference.dbo.Clients AS Clients ON MegaOrders.ClientID = Clients.ClientID
                WHERE MegaOrderID NOT IN (SELECT MegaOrderID FROM CabFurDispatch)
                AND MegaOrderID IN (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM DecorOrders " + filter + " ))";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(dt) > 0)
                {
                    int[] MegaOrders = new int[dt.Rows.Count];

                    for (int i = 0; i < dt.Rows.Count; i++)
                        MegaOrders[i] = Convert.ToInt32(Convert.ToInt32(dt.Rows[i]["MegaOrderID"]));
                    FilterCabFurOrders(MegaOrders);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dt.Rows[i]["Weight"] = GetWeight(Convert.ToInt32(dt.Rows[i]["MegaOrderID"]));
                    }
                }
            }

            foreach (DataRow dr in dt.Rows)
                MegaOrdersDT.Rows.Add(dr.ItemArray);
            dt.Dispose();

            FillPackagesInfo();
            MegaOrdersBS.MoveFirst();
        }

        public void FilterAssembleByDate(DateTime Date)
        {
            DataTable dt = MegaOrdersDT.Clone();

            string SelectCommand = @"SELECT MegaOrders.MegaOrderID, MegaOrders.ClientID, MegaOrders.OrderNumber, MegaOrders.OrderDate, Clients.ClientName FROM MegaOrders
                INNER JOIN CabFurDispatch ON MegaOrders.MegaOrderID = CabFurDispatch.MegaOrderID
                LEFT JOIN infiniu2_marketingreference.dbo.Clients AS Clients ON MegaOrders.ClientID = Clients.ClientID
                WHERE CAST(CabFurDispatch.PrepareDateTime AS DATE) = '" + Date.ToString("yyyy-MM-dd") + "'";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(dt) > 0)
                {
                    int[] MegaOrders = new int[dt.Rows.Count];

                    for (int i = 0; i < dt.Rows.Count; i++)
                        MegaOrders[i] = Convert.ToInt32(Convert.ToInt32(dt.Rows[i]["MegaOrderID"]));
                    FilterCabFurOrders(MegaOrders);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dt.Rows[i]["Weight"] = GetWeight(Convert.ToInt32(dt.Rows[i]["MegaOrderID"]));
                    }
                }
            }

            foreach (DataRow dr in dt.Rows)
                MegaOrdersDT.Rows.Add(dr.ItemArray);
            dt.Dispose();

            FillPackagesInfo();
        }

        public void FilterMainOrders(int MegaOrderID)
        {
            CabFurniturePackages.Clear();
            string SelectCommand = "SELECT MainOrders.MegaOrderID, MegaOrders.ClientID, MegaOrders.OrderNumber, MainOrders.FactoryID, MainOrders.MainOrderID " +
                "FROM MainOrders INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID WHERE MainOrders.MegaOrderID=" + MegaOrderID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MainOrdersDT);
            }
        }

        public void MoveToAssembleDate(DateTime AssembleDate)
        {
            AssembleDatesBS.Position = AssembleDatesBS.Find("PrepareDateTime", AssembleDate);
        }

        public void MoveToMegaOrder(int MegaOrderID)
        {
            MegaOrdersBS.Position = MegaOrdersBS.Find("MegaOrderID", MegaOrderID);
        }
    }

}
