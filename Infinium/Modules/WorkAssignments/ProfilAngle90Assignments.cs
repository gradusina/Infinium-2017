﻿using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.WorkAssignments
{
    public class ProfilAngle90Assignments : IFirstProfilName, IColorName, IInsetTypeName, IInsetColorName
    {
        #region Fields

        private DataTable Additions1Marsel4DT;
        private DataTable Additions1Marsel5DT;
        private DataTable Additions2Marsel4DT;
        private DataTable Additions2Marsel5DT;
        private DataTable Additions3Marsel4DT;
        private DataTable Additions3Marsel5DT;
        private DataTable ArchDecorOrdersDT;
        private DataTable AssemblyDT;
        private DataTable BagetWithAngelOrdersDT;
        private DataTable BagetWithAngleAssemblyDT;
        private bool bImpostMargin = false;
        private int ColorType = 0;
        private DataTable ComecDT;
        private DataTable RapidTechnoInsetDT;
        private DateTime CurrentDate;
        private DataTable DecorAssemblyDT;
        private DataTable DecorDT;
        private DataTable DecorParametersDT;
        private DataTable DeyingDT;
        private DataTable DistMainOrdersDT;
        private readonly FileManager FM = new FileManager();
        private DataTable FrameColorsDataTable = null;
        private DataTable FrontsDataTable = null;
        private ArrayList FrontsID;
        private DataTable FrontsOrdersDT;
        private int FrontType = 0;
        private DataTable GridsDecorOrdersDT;
        private DataTable InsetColorsDataTable = null;
        private DataTable InsetDT;
        private DataTable InsetTypesDataTable = null;
        private DataTable Jersy110GridsDT;
        private DataTable Jersy110OrdersDT;
        private DataTable Jersy110SimpleDT;
        private DataTable Jersy110VitrinaDT;
        private DataTable Marsel1GridsDT;
        private DataTable Marsel1OrdersDT;
        private DataTable Marsel1SimpleDT;
        private DataTable Marsel1VitrinaDT;
        private DataTable Marsel3GridsDT;
        private DataTable Marsel3OrdersDT;
        private DataTable Marsel3SimpleDT;
        private DataTable Marsel3VitrinaDT;
        private DataTable Marsel4GridsDT;
        private DataTable Marsel4OrdersDT;
        private DataTable Marsel4SimpleDT;
        private DataTable Marsel4VitrinaDT;
        private DataTable Marsel5GridsDT;
        private DataTable Marsel5OrdersDT;
        private DataTable Marsel5SimpleDT;
        private DataTable Marsel5VitrinaDT;
        private DataTable MonteGridsDT;
        private DataTable MonteOrdersDT;
        private DataTable MonteSimpleDT;
        private DataTable MonteVitrinaDT;
        private DataTable NotArchDecorOrdersDT;
        private DataTable PatinaDataTable = null;
        private DataTable pFlorencGridsDT;
        private DataTable pFlorencOrdersDT;
        private DataTable pFlorencSimpleDT;
        private DataTable pFlorencVitrinaDT;
        private DataTable pFoxGridsDT;
        private DataTable pFoxOrdersDT;
        private DataTable pFoxSimpleDT;
        private DataTable pFoxVitrinaDT;
        private DataTable PortoGridsDT;
        private DataTable PortoOrdersDT;
        private DataTable PortoSimpleDT;
        private DataTable PortoVitrinaDT;
        private DataTable PR1GridsDT;
        private DataTable PR1OrdersDT;
        private DataTable PR1SimpleDT;
        private DataTable PR1VitrinaDT;
        private DataTable PR3GridsDT;
        private DataTable PR3LuxDT;
        private DataTable PR3OrdersDT;
        private DataTable PR3SimpleDT;
        private DataTable PR3VitrinaDT;
        private DataTable ProfileNamesDT;
        private DataTable PRU8GridsDT;
        private DataTable PRU8OrdersDT;
        private DataTable PRU8SimpleDT;
        private DataTable PRU8VitrinaDT;
        private DataTable RapidDT;
        private DataTable ShervudGridsDT;
        private DataTable ShervudOrdersDT;
        private DataTable ShervudSimpleDT;
        private DataTable ShervudVitrinaDT;
        private DataTable StandardImpostDataTable;
        private DataTable StemasDT;
        private DataTable SummOrdersDT;
        private DataTable Techno1GridsDT;
        private DataTable Techno1LuxDT;
        private DataTable Techno1MegaDT;
        private DataTable Techno1OrdersDT;
        private DataTable Techno1SimpleDT;
        private DataTable Techno1VitrinaDT;
        private DataTable Techno2GridsDT;
        private DataTable Techno2LuxDT;
        private DataTable Techno2MegaDT;
        private DataTable Techno2OrdersDT;
        private DataTable Techno2SimpleDT;
        private DataTable Techno2VitrinaDT;
        private DataTable Techno4GridsDT;
        private DataTable Techno4LuxDT;
        private DataTable Techno4MegaDT;
        private DataTable Techno4OrdersDT;
        private DataTable Techno4SimpleDT;
        private DataTable Techno4VitrinaDT;
        private DataTable Techno5GridsDT;
        private DataTable Techno5LuxDT;
        private DataTable Techno5OrdersDT;
        private DataTable Techno5SimpleDT;
        private DataTable Techno5VitrinaDT;
        private DataTable TotalInfoDT;
        private DataTable WidthMegaInsetsDT;

        #endregion Fields

        #region Properties

        public ArrayList GetFrontsID
        {
            set => FrontsID = value;
        }

        #endregion Properties

        #region Constructors

        public ProfilAngle90Assignments()
        {
            Create();
            Fill();
        }

        #endregion Constructors

        #region Methods

        public void ClearOrders()
        {
            FrontsID.Clear();

            BagetWithAngelOrdersDT.Clear();
            NotArchDecorOrdersDT.Clear();
            ArchDecorOrdersDT.Clear();
            GridsDecorOrdersDT.Clear();

            Marsel1OrdersDT.Clear();
            Marsel5OrdersDT.Clear();
            PortoOrdersDT.Clear();
            MonteOrdersDT.Clear();
            Marsel3OrdersDT.Clear();
            Marsel4OrdersDT.Clear();
            Jersy110OrdersDT.Clear();
            ShervudOrdersDT.Clear();
            Techno1OrdersDT.Clear();
            Techno2OrdersDT.Clear();
            Techno4OrdersDT.Clear();
            pFoxOrdersDT.Clear();
            pFlorencOrdersDT.Clear();
            Techno5OrdersDT.Clear();
            PR1OrdersDT.Clear();
            PR3OrdersDT.Clear();
            PRU8OrdersDT.Clear();

            PR1SimpleDT.Clear();
            PR1VitrinaDT.Clear();
            PR1GridsDT.Clear();

            Marsel1SimpleDT.Clear();
            Marsel5SimpleDT.Clear();
            PortoSimpleDT.Clear();
            MonteSimpleDT.Clear();
            Marsel3SimpleDT.Clear();
            Marsel4SimpleDT.Clear();
            Jersy110SimpleDT.Clear();
            ShervudSimpleDT.Clear();
            Techno1SimpleDT.Clear();
            Techno2SimpleDT.Clear();
            Techno4SimpleDT.Clear();
            pFoxSimpleDT.Clear();
            pFlorencSimpleDT.Clear();
            Techno5SimpleDT.Clear();
            PR3SimpleDT.Clear();
            PRU8SimpleDT.Clear();

            Marsel1VitrinaDT.Clear();
            Marsel5VitrinaDT.Clear();
            PortoVitrinaDT.Clear();
            MonteVitrinaDT.Clear();
            Marsel3VitrinaDT.Clear();
            Marsel4VitrinaDT.Clear();
            Jersy110VitrinaDT.Clear();
            ShervudVitrinaDT.Clear();
            Techno1VitrinaDT.Clear();
            Techno2VitrinaDT.Clear();
            pFoxVitrinaDT.Clear();
            pFlorencVitrinaDT.Clear();
            Techno4VitrinaDT.Clear();
            Techno5VitrinaDT.Clear();
            PR3VitrinaDT.Clear();
            PRU8VitrinaDT.Clear();

            Marsel1GridsDT.Clear();
            Marsel5GridsDT.Clear();
            PortoGridsDT.Clear();
            MonteGridsDT.Clear();
            Marsel3GridsDT.Clear();
            Marsel4GridsDT.Clear();
            Jersy110GridsDT.Clear();
            Techno1GridsDT.Clear();
            Techno2GridsDT.Clear();
            Techno4GridsDT.Clear();
            pFoxGridsDT.Clear();
            pFlorencGridsDT.Clear();
            ShervudGridsDT.Clear();
            Techno5GridsDT.Clear();
            PR3GridsDT.Clear();
            PRU8GridsDT.Clear();

            Techno1LuxDT.Clear();
            Techno2LuxDT.Clear();
            Techno4LuxDT.Clear();
            Techno5LuxDT.Clear();
            PR3LuxDT.Clear();

            Techno1MegaDT.Clear();
            Techno2MegaDT.Clear();
            Techno4MegaDT.Clear();
        }

        public void CreateExcel(int WorkAssignmentID, int FactoryID, string BatchName, string ClientName, ref string sSourceFileName)
        {
            GetCurrentDate();

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            #region Create fonts and styles

            HSSFFont Calibri11F = hssfworkbook.CreateFont();
            Calibri11F.FontHeightInPoints = 11;
            Calibri11F.FontName = "Calibri";

            HSSFFont CalibriBold11F = hssfworkbook.CreateFont();
            CalibriBold11F.FontHeightInPoints = 11;
            CalibriBold11F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold11F.FontName = "Calibri";

            HSSFFont CalibriBold15F = hssfworkbook.CreateFont();
            CalibriBold15F.FontHeightInPoints = 15;
            CalibriBold15F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold15F.FontName = "Calibri";

            HSSFFont Serif8F = hssfworkbook.CreateFont();
            Serif8F.FontHeightInPoints = 8;
            Serif8F.FontName = "MS Sans Serif";

            HSSFFont Serif10F = hssfworkbook.CreateFont();
            Serif10F.FontHeightInPoints = 10;
            Serif10F.FontName = "MS Sans Serif";

            HSSFFont SerifBold10F = hssfworkbook.CreateFont();
            SerifBold10F.FontHeightInPoints = 10;
            SerifBold10F.FontName = "MS Sans Serif";
            SerifBold10F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;

            HSSFCellStyle Calibri11CS = hssfworkbook.CreateCellStyle();
            Calibri11CS.SetFont(Calibri11F);

            HSSFCellStyle CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.SetFont(CalibriBold11F);

            HSSFCellStyle CalibriBold15CS = hssfworkbook.CreateCellStyle();
            CalibriBold15CS.SetFont(CalibriBold15F);

            HSSFCellStyle TableHeaderCS = hssfworkbook.CreateCellStyle();
            TableHeaderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.SetFont(Serif8F);

            HSSFCellStyle TableHeaderDecCS = hssfworkbook.CreateCellStyle();
            TableHeaderDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.SetFont(Serif8F);

            HSSFCellStyle TableContentCS = hssfworkbook.CreateCellStyle();
            TableContentCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableContentCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableContentCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableContentCS.RightBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableContentCS.TopBorderColor = HSSFColor.BLACK.index;
            TableContentCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountCS = hssfworkbook.CreateCellStyle();
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountDecCS = hssfworkbook.CreateCellStyle();
            TotamAmountDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            #endregion Create fonts and styles

            if (Marsel1OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Marsel1OrdersDT, ref Marsel1SimpleDT);
                GetVitrinaFronts(Marsel1OrdersDT, ref Marsel1VitrinaDT);
                GetGridFronts(Marsel1OrdersDT, ref Marsel1GridsDT);
            }
            if (Marsel5OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Marsel5OrdersDT, ref Marsel5SimpleDT);
                GetVitrinaFronts(Marsel5OrdersDT, ref Marsel5VitrinaDT);
                GetGridFronts(Marsel5OrdersDT, ref Marsel5GridsDT);
            }
            if (PortoOrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(PortoOrdersDT, ref PortoSimpleDT);
                GetVitrinaFronts(PortoOrdersDT, ref PortoVitrinaDT);
                GetGridFronts(PortoOrdersDT, ref PortoGridsDT);
            }
            if (MonteOrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(MonteOrdersDT, ref MonteSimpleDT);
                GetVitrinaFronts(MonteOrdersDT, ref MonteVitrinaDT);
                GetGridFronts(MonteOrdersDT, ref MonteGridsDT);
            }
            if (Marsel3OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Marsel3OrdersDT, ref Marsel3SimpleDT);
                GetVitrinaFronts(Marsel3OrdersDT, ref Marsel3VitrinaDT);
                GetGridFronts(Marsel3OrdersDT, ref Marsel3GridsDT);
            }
            if (Marsel4OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Marsel4OrdersDT, ref Marsel4SimpleDT);
                GetVitrinaFronts(Marsel4OrdersDT, ref Marsel4VitrinaDT);
                GetGridFronts(Marsel4OrdersDT, ref Marsel4GridsDT);
            }
            if (Jersy110OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Jersy110OrdersDT, ref Jersy110SimpleDT);
                GetVitrinaFronts(Jersy110OrdersDT, ref Jersy110VitrinaDT);
                GetGridFronts(Jersy110OrdersDT, ref Jersy110GridsDT);
            }
            if (Techno1OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Techno1OrdersDT, ref Techno1SimpleDT);
                GetVitrinaFronts(Techno1OrdersDT, ref Techno1VitrinaDT);
                GetGridFronts(Techno1OrdersDT, ref Techno1GridsDT);
                GetLuxFronts(Techno1OrdersDT, ref Techno1LuxDT);
                GetMegaFronts(Techno1OrdersDT, ref Techno1MegaDT);
            }
            if (ShervudOrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(ShervudOrdersDT, ref ShervudSimpleDT);
                GetVitrinaFronts(ShervudOrdersDT, ref ShervudVitrinaDT);
                GetGridFronts(ShervudOrdersDT, ref ShervudGridsDT);
            }
            if (Techno2OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Techno2OrdersDT, ref Techno2SimpleDT);
                GetVitrinaFronts(Techno2OrdersDT, ref Techno2VitrinaDT);
                GetGridFronts(Techno2OrdersDT, ref Techno2GridsDT);
                GetLuxFronts(Techno2OrdersDT, ref Techno2LuxDT);
                GetMegaFronts(Techno2OrdersDT, ref Techno2MegaDT);
            }
            if (Techno4OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Techno4OrdersDT, ref Techno4SimpleDT);
                GetVitrinaFronts(Techno4OrdersDT, ref Techno4VitrinaDT);
                GetGridFronts(Techno4OrdersDT, ref Techno4GridsDT);
                GetLuxFronts(Techno4OrdersDT, ref Techno4LuxDT);
                GetMegaFronts(Techno4OrdersDT, ref Techno4MegaDT);
            }
            if (pFoxOrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(pFoxOrdersDT, ref pFoxSimpleDT);
                GetVitrinaFronts(pFoxOrdersDT, ref pFoxVitrinaDT);
                GetGridFronts(pFoxOrdersDT, ref pFoxGridsDT);
            }
            if (pFlorencOrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(pFlorencOrdersDT, ref pFlorencSimpleDT);
                GetVitrinaFronts(pFlorencOrdersDT, ref pFlorencVitrinaDT);
                GetGridFronts(pFlorencOrdersDT, ref pFlorencGridsDT);
            }
            if (Techno5OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(Techno5OrdersDT, ref Techno5SimpleDT);
                GetVitrinaFronts(Techno5OrdersDT, ref Techno5VitrinaDT);
                GetGridFronts(Techno5OrdersDT, ref Techno5GridsDT);
                GetLuxFronts(Techno5OrdersDT, ref Techno5LuxDT);
            }

            if (Marsel1OrdersDT.Rows.Count == 0 && Marsel5OrdersDT.Rows.Count == 0 && PortoOrdersDT.Rows.Count == 0 && MonteOrdersDT.Rows.Count == 0 &&
                Marsel3OrdersDT.Rows.Count == 0 && Marsel4OrdersDT.Rows.Count == 0 &&
                Jersy110OrdersDT.Rows.Count == 0 && ShervudOrdersDT.Rows.Count == 0 &&
                Techno1OrdersDT.Rows.Count == 0 && Techno2OrdersDT.Rows.Count == 0 && Techno4OrdersDT.Rows.Count == 0 && pFoxOrdersDT.Rows.Count == 0 && pFlorencOrdersDT.Rows.Count == 0 &&
                Techno5OrdersDT.Rows.Count == 0)
                return;

            string DispatchDate = string.Empty;
            if (ClientName == "ЗОВ" || ClientName == "Маркетинг + ЗОВ")
            {
                string FrontsFilterString = "(SELECT MainOrderID FROM FrontsOrders WHERE FactoryID=1 AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel3) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel4) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Jersy110) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel5) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Porto) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Monte) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno1) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Shervud) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno2) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno4) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.pFox) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.pFlorenc) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno5) + "))";
                string SelectCommand = @"SELECT DispatchDate, MegaOrderID FROM MegaOrders
                    WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders
                    WHERE MainOrderID IN" + FrontsFilterString + " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";

                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                    ConnectionStrings.ZOVOrdersConnectionString))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0 && DT.Rows[0]["DispatchDate"] != DBNull.Value)
                            DispatchDate = Convert.ToDateTime(DT.Rows[0]["DispatchDate"]).ToString("dd.MM.yyyy");
                    }
                }
            }

            BagetWithAngleAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            TotalInfoToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName);

            StemasToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, false);

            RapidToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, false);

            AdditionsToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, ClientName);

            InsetToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, false);

            AssemblyToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, false);

            OrdersSummaryInfoToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, false, false);

            DeyingByMainOrderToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GetMainOrdersSummary(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, false, false, false);

            string FileName = WorkAssignmentID + " " + BatchName + "  Угол 90";
            string tempFolder = @"\\192.168.1.6\Public\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ПРОФИЛЬ\инфиниум\";
            //string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");

            string CurrentMonthName = DateTime.Now.ToString("MMMM");
            tempFolder = Path.Combine(tempFolder, CurrentMonthName);
            if (!(Directory.Exists(tempFolder)))
            {
                Directory.CreateDirectory(tempFolder);
            }
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            sw.Stop();
            System.Diagnostics.Process.Start(file.FullName);
        }

        public void CreateExcelPR1(int WorkAssignmentID, int FactoryID, string BatchName, string ClientName, ref string sSourceFileName)
        {
            GetCurrentDate();

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            #region Create fonts and styles

            HSSFFont Calibri11F = hssfworkbook.CreateFont();
            Calibri11F.FontHeightInPoints = 11;
            Calibri11F.FontName = "Calibri";

            HSSFFont CalibriBold11F = hssfworkbook.CreateFont();
            CalibriBold11F.FontHeightInPoints = 11;
            CalibriBold11F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold11F.FontName = "Calibri";

            HSSFFont CalibriBold15F = hssfworkbook.CreateFont();
            CalibriBold15F.FontHeightInPoints = 15;
            CalibriBold15F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold15F.FontName = "Calibri";

            HSSFFont Serif8F = hssfworkbook.CreateFont();
            Serif8F.FontHeightInPoints = 8;
            Serif8F.FontName = "MS Sans Serif";

            HSSFFont Serif10F = hssfworkbook.CreateFont();
            Serif10F.FontHeightInPoints = 10;
            Serif10F.FontName = "MS Sans Serif";

            HSSFFont SerifBold10F = hssfworkbook.CreateFont();
            SerifBold10F.FontHeightInPoints = 10;
            SerifBold10F.FontName = "MS Sans Serif";
            SerifBold10F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;

            HSSFCellStyle Calibri11CS = hssfworkbook.CreateCellStyle();
            Calibri11CS.SetFont(Calibri11F);

            HSSFCellStyle CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.SetFont(CalibriBold11F);

            HSSFCellStyle CalibriBold15CS = hssfworkbook.CreateCellStyle();
            CalibriBold15CS.SetFont(CalibriBold15F);

            HSSFCellStyle TableHeaderCS = hssfworkbook.CreateCellStyle();
            TableHeaderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.SetFont(Serif8F);

            HSSFCellStyle TableHeaderDecCS = hssfworkbook.CreateCellStyle();
            TableHeaderDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.SetFont(Serif8F);

            HSSFCellStyle TableContentCS = hssfworkbook.CreateCellStyle();
            TableContentCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableContentCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableContentCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableContentCS.RightBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableContentCS.TopBorderColor = HSSFColor.BLACK.index;
            TableContentCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountCS = hssfworkbook.CreateCellStyle();
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountDecCS = hssfworkbook.CreateCellStyle();
            TotamAmountDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            #endregion Create fonts and styles

            PR1SimpleDT.Clear();

            PR1VitrinaDT.Clear();

            PR1GridsDT.Clear();

            Marsel1SimpleDT.Clear();
            Marsel5SimpleDT.Clear();
            PortoSimpleDT.Clear();
            MonteSimpleDT.Clear();
            Marsel3SimpleDT.Clear();
            Marsel4SimpleDT.Clear();
            Jersy110SimpleDT.Clear();
            ShervudSimpleDT.Clear();
            Techno1SimpleDT.Clear();
            Techno2SimpleDT.Clear();
            Techno4SimpleDT.Clear();
            pFoxSimpleDT.Clear();
            pFlorencSimpleDT.Clear();
            Techno5SimpleDT.Clear();
            PR3SimpleDT.Clear();
            PRU8SimpleDT.Clear();

            Marsel1VitrinaDT.Clear();
            Marsel5VitrinaDT.Clear();
            PortoVitrinaDT.Clear();
            MonteVitrinaDT.Clear();
            Marsel3VitrinaDT.Clear();
            Marsel4VitrinaDT.Clear();
            Jersy110VitrinaDT.Clear();
            ShervudVitrinaDT.Clear();
            Techno1VitrinaDT.Clear();
            Techno2VitrinaDT.Clear();
            Techno4VitrinaDT.Clear();
            pFoxVitrinaDT.Clear();
            pFlorencVitrinaDT.Clear();
            Techno5VitrinaDT.Clear();
            PR3VitrinaDT.Clear();
            PRU8VitrinaDT.Clear();

            Marsel1GridsDT.Clear();
            Marsel5GridsDT.Clear();
            PortoGridsDT.Clear();
            MonteGridsDT.Clear();
            Marsel3GridsDT.Clear();
            Marsel4GridsDT.Clear();
            Jersy110GridsDT.Clear();
            Techno1GridsDT.Clear();
            Techno2GridsDT.Clear();
            Techno4GridsDT.Clear();
            pFoxGridsDT.Clear();
            pFlorencGridsDT.Clear();
            ShervudGridsDT.Clear();
            Techno5GridsDT.Clear();
            PR3GridsDT.Clear();
            PRU8GridsDT.Clear();

            Techno1LuxDT.Clear();
            Techno2LuxDT.Clear();
            Techno4LuxDT.Clear();
            Techno5LuxDT.Clear();
            PR3LuxDT.Clear();

            Techno1MegaDT.Clear();
            Techno2MegaDT.Clear();
            Techno4MegaDT.Clear();

            if (PR1OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(PR1OrdersDT, ref PR1SimpleDT);
                GetVitrinaFronts(PR1OrdersDT, ref PR1VitrinaDT);
                GetGridFronts(PR1OrdersDT, ref PR1GridsDT);
            }

            if (PR1OrdersDT.Rows.Count == 0)
                return;

            string DispatchDate = string.Empty;
            if (ClientName == "ЗОВ" || ClientName == "Маркетинг + ЗОВ")
            {
                string FrontsFilterString = "(SELECT MainOrderID FROM FrontsOrders WHERE FactoryID=1 AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR1) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR2) + "))";
                string SelectCommand = @"SELECT DispatchDate, MegaOrderID FROM MegaOrders
                    WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders
                    WHERE MainOrderID IN" + FrontsFilterString + " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";

                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                    ConnectionStrings.ZOVOrdersConnectionString))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0 && DT.Rows[0]["DispatchDate"] != DBNull.Value)
                            DispatchDate = Convert.ToDateTime(DT.Rows[0]["DispatchDate"]).ToString("dd.MM.yyyy");
                    }
                }
            }

            BagetWithAngleAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            StemasToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, true, false, false);

            RapidToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, true, false, false);

            InsetToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, true, false, false);

            AssemblyToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, true, false, false);

            OrdersSummaryInfoToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, true, true, false, false);

            DeyingPR1ByMainOrderToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GetMainOrdersSummary(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, true, false, false);

            //string FileName = "№" + WorkAssignmentID + " " + BatchName;
            //string tempFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            string FileName = WorkAssignmentID + " " + BatchName + "  ПР-1 и ПР-2";
            string tempFolder = @"\\192.168.1.6\Public\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ПРОФИЛЬ\инфиниум\";
            //string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");

            string CurrentMonthName = DateTime.Now.ToString("MMMM");
            tempFolder = Path.Combine(tempFolder, CurrentMonthName);
            if (!(Directory.Exists(tempFolder)))
            {
                Directory.CreateDirectory(tempFolder);
            }
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            sw.Stop();
            System.Diagnostics.Process.Start(file.FullName);
        }

        public void CreateExcelPR3(int WorkAssignmentID, int FactoryID, string BatchName, string ClientName, ref string sSourceFileName)
        {
            GetCurrentDate();

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            #region Create fonts and styles

            HSSFFont Calibri11F = hssfworkbook.CreateFont();
            Calibri11F.FontHeightInPoints = 11;
            Calibri11F.FontName = "Calibri";

            HSSFFont CalibriBold11F = hssfworkbook.CreateFont();
            CalibriBold11F.FontHeightInPoints = 11;
            CalibriBold11F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold11F.FontName = "Calibri";

            HSSFFont CalibriBold15F = hssfworkbook.CreateFont();
            CalibriBold15F.FontHeightInPoints = 15;
            CalibriBold15F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold15F.FontName = "Calibri";

            HSSFFont Serif8F = hssfworkbook.CreateFont();
            Serif8F.FontHeightInPoints = 8;
            Serif8F.FontName = "MS Sans Serif";

            HSSFFont Serif10F = hssfworkbook.CreateFont();
            Serif10F.FontHeightInPoints = 10;
            Serif10F.FontName = "MS Sans Serif";

            HSSFFont SerifBold10F = hssfworkbook.CreateFont();
            SerifBold10F.FontHeightInPoints = 10;
            SerifBold10F.FontName = "MS Sans Serif";
            SerifBold10F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;

            HSSFCellStyle Calibri11CS = hssfworkbook.CreateCellStyle();
            Calibri11CS.SetFont(Calibri11F);

            HSSFCellStyle CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.SetFont(CalibriBold11F);

            HSSFCellStyle CalibriBold15CS = hssfworkbook.CreateCellStyle();
            CalibriBold15CS.SetFont(CalibriBold15F);

            HSSFCellStyle TableHeaderCS = hssfworkbook.CreateCellStyle();
            TableHeaderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.SetFont(Serif8F);

            HSSFCellStyle TableHeaderDecCS = hssfworkbook.CreateCellStyle();
            TableHeaderDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.SetFont(Serif8F);

            HSSFCellStyle TableContentCS = hssfworkbook.CreateCellStyle();
            TableContentCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableContentCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableContentCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableContentCS.RightBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableContentCS.TopBorderColor = HSSFColor.BLACK.index;
            TableContentCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountCS = hssfworkbook.CreateCellStyle();
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountDecCS = hssfworkbook.CreateCellStyle();
            TotamAmountDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            #endregion Create fonts and styles

            PR1SimpleDT.Clear();

            PR1VitrinaDT.Clear();

            PR1GridsDT.Clear();

            Marsel1SimpleDT.Clear();
            Marsel5SimpleDT.Clear();
            PortoSimpleDT.Clear();
            MonteSimpleDT.Clear();
            Marsel3SimpleDT.Clear();
            Marsel4SimpleDT.Clear();
            Jersy110SimpleDT.Clear();
            ShervudSimpleDT.Clear();
            Techno1SimpleDT.Clear();
            Techno2SimpleDT.Clear();
            Techno4SimpleDT.Clear();
            pFoxSimpleDT.Clear();
            pFlorencSimpleDT.Clear();
            Techno5SimpleDT.Clear();
            PR3SimpleDT.Clear();
            PRU8SimpleDT.Clear();

            Marsel1VitrinaDT.Clear();
            Marsel5VitrinaDT.Clear();
            PortoVitrinaDT.Clear();
            MonteVitrinaDT.Clear();
            Marsel3VitrinaDT.Clear();
            Marsel4VitrinaDT.Clear();
            Jersy110VitrinaDT.Clear();
            ShervudVitrinaDT.Clear();
            Techno1VitrinaDT.Clear();
            Techno2VitrinaDT.Clear();
            Techno4VitrinaDT.Clear();
            pFoxVitrinaDT.Clear();
            pFlorencVitrinaDT.Clear();
            Techno5VitrinaDT.Clear();
            PR3VitrinaDT.Clear();
            PRU8VitrinaDT.Clear();

            Marsel1GridsDT.Clear();
            Marsel5GridsDT.Clear();
            PortoGridsDT.Clear();
            MonteGridsDT.Clear();
            Marsel3GridsDT.Clear();
            Marsel4GridsDT.Clear();
            Jersy110GridsDT.Clear();
            Techno1GridsDT.Clear();
            Techno2GridsDT.Clear();
            Techno4GridsDT.Clear();
            pFoxGridsDT.Clear();
            pFlorencGridsDT.Clear();
            ShervudGridsDT.Clear();
            Techno5GridsDT.Clear();
            PR3GridsDT.Clear();
            PRU8GridsDT.Clear();

            Techno1LuxDT.Clear();
            Techno2LuxDT.Clear();
            Techno4LuxDT.Clear();
            Techno5LuxDT.Clear();
            PR3LuxDT.Clear();

            Techno1MegaDT.Clear();
            Techno2MegaDT.Clear();
            Techno4MegaDT.Clear();

            if (PR3OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(PR3OrdersDT, ref PR3SimpleDT);
                GetVitrinaFronts(PR3OrdersDT, ref PR3VitrinaDT);
                GetGridFronts(PR3OrdersDT, ref PR3GridsDT);
                GetLuxFronts(PR3OrdersDT, ref PR3LuxDT);
            }

            if (PR3OrdersDT.Rows.Count == 0)
                return;

            string DispatchDate = string.Empty;
            if (ClientName == "ЗОВ" || ClientName == "Маркетинг + ЗОВ")
            {
                string FrontsFilterString = "(SELECT MainOrderID FROM FrontsOrders WHERE FactoryID=1 AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR3) + "))";
                string SelectCommand = @"SELECT DispatchDate, MegaOrderID FROM MegaOrders
                    WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders
                    WHERE MainOrderID IN" + FrontsFilterString + " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";

                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                    ConnectionStrings.ZOVOrdersConnectionString))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0 && DT.Rows[0]["DispatchDate"] != DBNull.Value)
                            DispatchDate = Convert.ToDateTime(DT.Rows[0]["DispatchDate"]).ToString("dd.MM.yyyy");
                    }
                }
            }

            BagetWithAngleAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            StemasToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, true, false);

            RapidToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, true, false);

            InsetToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, true, false);

            AssemblyToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, true, false);

            OrdersSummaryInfoToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, true, false);

            DeyingPR3ByMainOrderToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GetMainOrdersSummary(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, false, true, false);

            string FileName = WorkAssignmentID + " " + BatchName + "  ПР-3";
            string tempFolder = @"\\192.168.1.6\Public\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ПРОФИЛЬ\инфиниум\";
            //string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");

            string CurrentMonthName = DateTime.Now.ToString("MMMM");
            tempFolder = Path.Combine(tempFolder, CurrentMonthName);
            if (!(Directory.Exists(tempFolder)))
            {
                Directory.CreateDirectory(tempFolder);
            }
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            sw.Stop();
            System.Diagnostics.Process.Start(file.FullName);

            //string sSourceFolder = System.Environment.GetEnvironmentVariable("TEMP");
            //string sFolderPath = "Общие файлы/Производство/Задания в работу";
            //string sDestFolder = Configs.DocumentsPath + sFolderPath;
            //sSourceFileName = GetFileName(sDestFolder, BatchName);

            //FileInfo file = new FileInfo(sSourceFolder + @"\" + sSourceFileName);
            //FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            //hssfworkbook.Write(NewFile);
            //NewFile.Close();
        }

        public void CreateExcelPRU8(int WorkAssignmentID, int FactoryID, string BatchName, string ClientName, ref string sSourceFileName)
        {
            GetCurrentDate();

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            #region Create fonts and styles

            HSSFFont Calibri11F = hssfworkbook.CreateFont();
            Calibri11F.FontHeightInPoints = 11;
            Calibri11F.FontName = "Calibri";

            HSSFFont CalibriBold11F = hssfworkbook.CreateFont();
            CalibriBold11F.FontHeightInPoints = 11;
            CalibriBold11F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold11F.FontName = "Calibri";

            HSSFFont CalibriBold15F = hssfworkbook.CreateFont();
            CalibriBold15F.FontHeightInPoints = 15;
            CalibriBold15F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold15F.FontName = "Calibri";

            HSSFFont Serif8F = hssfworkbook.CreateFont();
            Serif8F.FontHeightInPoints = 8;
            Serif8F.FontName = "MS Sans Serif";

            HSSFFont Serif10F = hssfworkbook.CreateFont();
            Serif10F.FontHeightInPoints = 10;
            Serif10F.FontName = "MS Sans Serif";

            HSSFFont SerifBold10F = hssfworkbook.CreateFont();
            SerifBold10F.FontHeightInPoints = 10;
            SerifBold10F.FontName = "MS Sans Serif";
            SerifBold10F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;

            HSSFCellStyle Calibri11CS = hssfworkbook.CreateCellStyle();
            Calibri11CS.SetFont(Calibri11F);

            HSSFCellStyle CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.SetFont(CalibriBold11F);

            HSSFCellStyle CalibriBold15CS = hssfworkbook.CreateCellStyle();
            CalibriBold15CS.SetFont(CalibriBold15F);

            HSSFCellStyle TableHeaderCS = hssfworkbook.CreateCellStyle();
            TableHeaderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.SetFont(Serif8F);

            HSSFCellStyle TableHeaderDecCS = hssfworkbook.CreateCellStyle();
            TableHeaderDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.SetFont(Serif8F);

            HSSFCellStyle TableContentCS = hssfworkbook.CreateCellStyle();
            TableContentCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableContentCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableContentCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableContentCS.RightBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableContentCS.TopBorderColor = HSSFColor.BLACK.index;
            TableContentCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountCS = hssfworkbook.CreateCellStyle();
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountDecCS = hssfworkbook.CreateCellStyle();
            TotamAmountDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            #endregion Create fonts and styles

            PR1SimpleDT.Clear();

            PR1VitrinaDT.Clear();

            PR1GridsDT.Clear();

            Marsel1SimpleDT.Clear();
            Marsel5SimpleDT.Clear();
            PortoSimpleDT.Clear();
            MonteSimpleDT.Clear();
            Marsel3SimpleDT.Clear();
            Marsel4SimpleDT.Clear();
            Jersy110SimpleDT.Clear();
            ShervudSimpleDT.Clear();
            Techno1SimpleDT.Clear();
            Techno2SimpleDT.Clear();
            Techno4SimpleDT.Clear();
            pFoxSimpleDT.Clear();
            pFlorencSimpleDT.Clear();
            Techno5SimpleDT.Clear();
            PR3SimpleDT.Clear();
            PRU8SimpleDT.Clear();

            Marsel1VitrinaDT.Clear();
            Marsel5VitrinaDT.Clear();
            PortoVitrinaDT.Clear();
            MonteVitrinaDT.Clear();
            Marsel3VitrinaDT.Clear();
            Marsel4VitrinaDT.Clear();
            Jersy110VitrinaDT.Clear();
            ShervudVitrinaDT.Clear();
            Techno1VitrinaDT.Clear();
            Techno2VitrinaDT.Clear();
            Techno4VitrinaDT.Clear();
            pFoxVitrinaDT.Clear();
            pFlorencVitrinaDT.Clear();
            Techno5VitrinaDT.Clear();
            PR3VitrinaDT.Clear();
            PRU8VitrinaDT.Clear();

            Marsel1GridsDT.Clear();
            Marsel5GridsDT.Clear();
            PortoGridsDT.Clear();
            MonteGridsDT.Clear();
            Marsel3GridsDT.Clear();
            Marsel4GridsDT.Clear();
            Jersy110GridsDT.Clear();
            Techno1GridsDT.Clear();
            Techno2GridsDT.Clear();
            Techno4GridsDT.Clear();
            pFoxGridsDT.Clear();
            pFlorencGridsDT.Clear();
            ShervudGridsDT.Clear();
            Techno5GridsDT.Clear();
            PR3GridsDT.Clear();
            PRU8GridsDT.Clear();

            Techno1LuxDT.Clear();
            Techno2LuxDT.Clear();
            Techno4LuxDT.Clear();
            Techno5LuxDT.Clear();
            PR3LuxDT.Clear();

            Techno1MegaDT.Clear();
            Techno2MegaDT.Clear();
            Techno4MegaDT.Clear();

            if (PRU8OrdersDT.Rows.Count > 0)
            {
                GetSimpleFronts(PRU8OrdersDT, ref PRU8SimpleDT);
                GetVitrinaFronts(PRU8OrdersDT, ref PRU8VitrinaDT);
                GetGridFronts(PRU8OrdersDT, ref PRU8GridsDT);
            }

            if (PRU8OrdersDT.Rows.Count == 0)
                return;

            string DispatchDate = string.Empty;
            if (ClientName == "ЗОВ" || ClientName == "Маркетинг + ЗОВ")
            {
                string FrontsFilterString = "(SELECT MainOrderID FROM FrontsOrders WHERE FactoryID=1 AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PRU8) + "))";
                string SelectCommand = @"SELECT DispatchDate, MegaOrderID FROM MegaOrders
                    WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders
                    WHERE MainOrderID IN" + FrontsFilterString + " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";
                
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                    ConnectionStrings.ZOVOrdersConnectionString))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0 && DT.Rows[0]["DispatchDate"] != DBNull.Value)
                            DispatchDate = Convert.ToDateTime(DT.Rows[0]["DispatchDate"]).ToString("dd.MM.yyyy");
                    }
                }
            }

            BagetWithAngleAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            StemasToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, true);

            RapidToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, true);

            InsetToExcel(ref hssfworkbook,
                CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, true);

            AssemblyToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, true);

            OrdersSummaryInfoToExcel(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, false, false, false, true);

            DeyingPRU8ByMainOrderToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GetMainOrdersSummary(ref hssfworkbook,
               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, false, false, true);

            //string FileName = "№" + WorkAssignmentID + " " + BatchName;

            //string tempFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            string FileName = WorkAssignmentID + " " + BatchName + "  ПРУ-8";
            string tempFolder = @"\\192.168.1.6\Public\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ПРОФИЛЬ\инфиниум\";
            //string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");

            string CurrentMonthName = DateTime.Now.ToString("MMMM");
            tempFolder = Path.Combine(tempFolder, CurrentMonthName); if (!(Directory.Exists(tempFolder)))
            {
                Directory.CreateDirectory(tempFolder);
            }
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            sw.Stop();
            System.Diagnostics.Process.Start(file.FullName);
        }

        public string GetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            try
            {
                DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
                ColorName = Rows[0]["ColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public string GetFrontName(int FrontID)
        {
            string FrontName = "";
            try
            {
                DataRow[] Rows = FrontsDataTable.Select("FrontID = " + FrontID);
                FrontName = Rows[0]["FrontName"].ToString();
            }
            catch
            {
                return "";
            }
            return FrontName;
        }

        public string GetInsetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            try
            {
                DataRow[] Rows = InsetColorsDataTable.Select("InsetColorID = " + ColorID);
                ColorName = Rows[0]["InsetColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public string GetInsetTypeName(int InsetTypeID)
        {
            string InsetType = string.Empty;
            try
            {
                DataRow[] Rows = InsetTypesDataTable.Select("InsetTypeID = " + InsetTypeID);
                InsetType = Rows[0]["InsetTypeName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return InsetType;
        }

        public bool GetOrders(int WorkAssignmentID, int FactoryID)
        {
            bImpostMargin = false;
            ProfileNamesDT.Clear();
            //InsetTypeNamesDT.Clear();
            for (int i = 0; i < FrontsID.Count; i++)
            {
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Marsel1))
                {
                    GetFrontsOrders(ref Marsel1OrdersDT, WorkAssignmentID, FactoryID, Fronts.Marsel1);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Marsel1);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Marsel1);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Marsel5))
                {
                    GetFrontsOrders(ref Marsel5OrdersDT, WorkAssignmentID, FactoryID, Fronts.Marsel5);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Marsel5);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Marsel5);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Porto))
                {
                    GetFrontsOrders(ref PortoOrdersDT, WorkAssignmentID, FactoryID, Fronts.Porto);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Marsel5);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Porto);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Monte))
                {
                    GetFrontsOrders(ref MonteOrdersDT, WorkAssignmentID, FactoryID, Fronts.Monte);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Monte);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Monte);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Marsel3))
                {
                    GetMarselOrders(ref Marsel3OrdersDT, WorkAssignmentID, FactoryID, Fronts.Marsel3);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Marsel3);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Marsel3);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Marsel4))
                {
                    GetMarselOrders(ref Marsel4OrdersDT, WorkAssignmentID, FactoryID, Fronts.Marsel4);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Marsel4);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Jersy110))
                {
                    GetFrontsOrders(ref Jersy110OrdersDT, WorkAssignmentID, FactoryID, Fronts.Jersy110);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Jersy110);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Techno1))
                {
                    GetFrontsOrders(ref Techno1OrdersDT, WorkAssignmentID, FactoryID, Fronts.Techno1);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno1);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno1);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Shervud))
                {
                    GetFrontsOrders(ref ShervudOrdersDT, WorkAssignmentID, FactoryID, Fronts.Shervud);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno1);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Shervud);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Techno2))
                {
                    GetFrontsOrders(ref Techno2OrdersDT, WorkAssignmentID, FactoryID, Fronts.Techno2);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno2);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno2);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Techno4))
                {
                    GetFrontsOrders(ref Techno4OrdersDT, WorkAssignmentID, FactoryID, Fronts.Techno4);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno4);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno4);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.pFox))
                {
                    GetFrontsOrders(ref pFoxOrdersDT, WorkAssignmentID, FactoryID, Fronts.pFox);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.pFox);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.pFox);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.pFlorenc))
                {
                    GetFrontsOrders(ref pFlorencOrdersDT, WorkAssignmentID, FactoryID, Fronts.pFlorenc);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.p1418);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.pFlorenc);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Techno5))
                {
                    GetFrontsOrders(ref Techno5OrdersDT, WorkAssignmentID, FactoryID, Fronts.Techno5);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno5);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Techno5);
                }
            }

            GetNotArchDecorOrders(ref NotArchDecorOrdersDT, WorkAssignmentID, FactoryID);
            GetBagetWithAngleOrders(ref BagetWithAngelOrdersDT, WorkAssignmentID, FactoryID);
            GetArchDecorOrders(ref ArchDecorOrdersDT, WorkAssignmentID, FactoryID);
            GetGridsDecorOrders(ref GridsDecorOrdersDT, WorkAssignmentID, FactoryID);

            if (Marsel1OrdersDT.Rows.Count == 0 && Marsel5OrdersDT.Rows.Count == 0 && PortoOrdersDT.Rows.Count == 0 && MonteOrdersDT.Rows.Count == 0 &&
                Marsel3OrdersDT.Rows.Count == 0 && Marsel4OrdersDT.Rows.Count == 0 &&
                Jersy110OrdersDT.Rows.Count == 0 &&
                ShervudOrdersDT.Rows.Count == 0 && Techno1OrdersDT.Rows.Count == 0 && Techno2OrdersDT.Rows.Count == 0 &&
                Techno4OrdersDT.Rows.Count == 0 && pFoxOrdersDT.Rows.Count == 0 && pFlorencOrdersDT.Rows.Count == 0 && Techno5OrdersDT.Rows.Count == 0 &&
                PR3OrdersDT.Rows.Count == 0 &&
                BagetWithAngelOrdersDT.Rows.Count == 0 && NotArchDecorOrdersDT.Rows.Count == 0 && ArchDecorOrdersDT.Rows.Count == 0 && GridsDecorOrdersDT.Rows.Count == 0)
                return false;
            else
                return true;
        }

        public bool GetPR1Orders(int WorkAssignmentID, int FactoryID)
        {
            ProfileNamesDT.Clear();
            //InsetTypeNamesDT.Clear();
            for (int i = 0; i < FrontsID.Count; i++)
            {
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.PR1))
                {
                    GetFrontsOrders(ref PR1OrdersDT, WorkAssignmentID, FactoryID, Fronts.PR1);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.PR1);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.PR1);
                }
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.PR2))
                {
                    GetFrontsOrders(ref PR1OrdersDT, WorkAssignmentID, FactoryID, Fronts.PR2);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.PR2);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.PR2);
                }
            }

            if (PR1OrdersDT.Rows.Count == 0)
                return false;
            else
                return true;
        }

        public bool GetPR3Orders(int WorkAssignmentID, int FactoryID)
        {
            ProfileNamesDT.Clear();
            //InsetTypeNamesDT.Clear();
            for (int i = 0; i < FrontsID.Count; i++)
            {
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.PR3))
                {
                    GetFrontsOrders(ref PR3OrdersDT, WorkAssignmentID, FactoryID, Fronts.PR3);
                    //GetInsetTypeNames(ref InsetTypeNamesDT, WorkAssignmentID, FactoryID, Fronts.PR3);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.PR3);
                }
            }

            if (PR3OrdersDT.Rows.Count == 0)
                return false;
            else
                return true;
        }

        public bool GetPRU8Orders(int WorkAssignmentID, int FactoryID)
        {
            ProfileNamesDT.Clear();
            //InsetTypeNamesDT.Clear();
            for (int i = 0; i < FrontsID.Count; i++)
            {
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.PRU8))
                {
                    GetFrontsOrders(ref PRU8OrdersDT, WorkAssignmentID, FactoryID, Fronts.PRU8);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.PRU8);
                }
            }

            if (PRU8OrdersDT.Rows.Count == 0)
                return false;
            else
                return true;
        }

        private void AdditionsSyngly(DataTable SourceDT, ref DataTable DestinationDT, int coef)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoInsetTypeID<>-1";
                DT1 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                DataRow[] Srows = SourceDT.Select("TechnoInsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoInsetColorID"]) + " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]));
                if (Srows.Count() > 0)
                {
                    int Count = 0;

                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    DataRow[] rows = DestinationDT.Select("TechnoInsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoInsetColorID"]) + " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]));
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["InsetType"] = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["TechnoInsetTypeID"]));
                        NewRow["InsetColor"] = GetInsetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoInsetColorID"]));
                        NewRow["Height"] = Convert.ToInt32(DT1.Rows[i]["Height"]);
                        NewRow["Width"] = Convert.ToInt32(DT1.Rows[i]["Width"]);
                        NewRow["Count"] = Count * coef;
                        NewRow["TechnoInsetTypeID"] = Convert.ToInt32(DT1.Rows[i]["TechnoInsetTypeID"]);
                        NewRow["TechnoInsetColorID"] = Convert.ToInt32(DT1.Rows[i]["TechnoInsetColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * coef;
                    }
                }
            }
            DT1.Dispose();
        }

        private void AdditionsToExcel(ref HSSFWorkbook hssfworkbook,
                                                                                                                                                                    HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName)
        {
            Additions1Marsel4DT.Clear();
            CollectAdditions(Marsel4SimpleDT, ref Additions1Marsel4DT, 1);

            Additions2Marsel4DT.Clear();
            CollectAdditions(Marsel4SimpleDT, ref Additions2Marsel4DT, 3);

            Additions3Marsel4DT.Clear();
            CollectAdditions(Marsel4SimpleDT, ref Additions3Marsel4DT, 1);

            Additions1Marsel5DT.Clear();
            CollectAdditions(Marsel5SimpleDT, ref Additions1Marsel5DT, 1);

            Additions2Marsel5DT.Clear();
            CollectAdditions(Marsel5SimpleDT, ref Additions2Marsel5DT, 3);

            Additions3Marsel5DT.Clear();
            CollectAdditions(Marsel5SimpleDT, ref Additions3Marsel5DT, 1);

            if (Additions1Marsel4DT.Rows.Count == 0 && Additions2Marsel4DT.Rows.Count == 0 && Additions3Marsel4DT.Rows.Count == 0 &&
                Additions1Marsel5DT.Rows.Count == 0 && Additions2Marsel5DT.Rows.Count == 0 && Additions3Marsel5DT.Rows.Count == 0)
                return;

            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Доп. задания");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 11 * 256);
            sheet1.SetColumnWidth(3, 11 * 256);
            sheet1.SetColumnWidth(4, 7 * 256);
            sheet1.SetColumnWidth(5, 11 * 256);

            DataTable DT = new DataTable();
            DataColumn Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));

            if (Additions1Marsel4DT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = Additions1Marsel4DT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(5);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Зарезание ремяк раскладки из ПФх-03", ref RowIndex);
                RowIndex++;
            }

            if (Additions2Marsel4DT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = Additions2Marsel4DT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(5);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Покраска торцов ремяк раскладки из ПФх-03 (2 слоя)", ref RowIndex);
                RowIndex++;
            }

            if (Additions3Marsel4DT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = Additions3Marsel4DT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(5);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Приклеивание ремяк раскладки из ПФх-03", ref RowIndex);
                RowIndex++;
            }

            if (Additions1Marsel5DT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = Additions1Marsel5DT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(5);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Comec", ref RowIndex);
                RowIndex++;
            }

            if (Additions2Marsel5DT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = Additions2Marsel5DT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(5);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Покраска торцов ремяк раскладки из ПФ-36 (2 слоя)", ref RowIndex);
                RowIndex++;
            }

            if (Additions3Marsel5DT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = Additions3Marsel5DT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(5);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Приклеивание ремяк раскладки из ПФ-36", ref RowIndex);
            }
        }

        private void AdditionsToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, ref int RowIndex)
        {
            HSSFCell cell = null;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота (фасада)");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина (фасада)");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["TechnoInsetColorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "TechnoInsetTypeID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1 && CType != Convert.ToInt32(DT.Rows[x + 1]["TechnoInsetColorID"]))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "TechnoInsetTypeID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "TechnoInsetTypeID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "TechnoInsetTypeID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);
        }

        private void AllInsets(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool NeedSwap, bool OrderASC, bool Impost)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                if (Impost)
                    DV.RowFilter = "TechnoColorID<>-1";
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    if (Impost)
                        DV.RowFilter = "TechnoColorID<>-1";
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        if (Impost)
                            DV.RowFilter = "TechnoColorID<>-1";
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            if (Impost)
                                DV.RowFilter = "TechnoColorID<>-1";
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - HeightMargin;
                            if (Impost)
                                Height = Convert.ToInt32(DT4.Rows[y]["Height"]) / 2 - HeightMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            //string InsetColor = GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            string InsetColor = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            int GroupID = Convert.ToInt32(InsetTypesDataTable.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]))[0]["GroupID"]);
                            if (GroupID == 7 || GroupID == 8)
                            {
                                //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");

                                if (Height > 100 && Width > 100)
                                    continue;
                            }
                            //if (Impost)
                            //{
                            //    InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");
                            //}

                            if (Width > 910)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                if (NeedSwap)
                                {
                                    NewRow["Width"] = Height;
                                    NewRow["Height"] = Width;
                                }
                                else
                                {
                                    NewRow["Height"] = Height;
                                    NewRow["Width"] = Width;
                                }
                                if (Impost)
                                    NewRow["Count"] = Count * 2;
                                else
                                    NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                if (Impost)
                                    rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                                else
                                    rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void AllInsetsMarsel3(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool NeedSwap, bool OrderASC, bool Impost)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                if (Impost)
                    DV.RowFilter = "TechnoColorID<>-1";
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    if (Impost)
                        DV.RowFilter = "TechnoColorID<>-1";
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        if (Impost)
                            DV.RowFilter = "TechnoColorID<>-1";
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            if (Impost)
                                DV.RowFilter = "TechnoColorID<>-1";
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - HeightMargin;
                            if (Impost)
                                Height = Convert.ToInt32(DT4.Rows[y]["Height"]) / 2 - HeightMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            string InsetColor = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            int GroupID = Convert.ToInt32(InsetTypesDataTable.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]))[0]["GroupID"]);
                            //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");
                            if (GroupID == 7 || GroupID == 8)
                            {
                                //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");

                                if (Height > 100 && Width > 100)
                                    continue;
                            }
                            //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");
                            //if (Impost)
                            //{
                            //    InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");
                            //}

                            if (Width > 910)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                if (NeedSwap)
                                {
                                    NewRow["Width"] = Height;
                                    NewRow["Height"] = Width;
                                }
                                else
                                {
                                    NewRow["Height"] = Height;
                                    NewRow["Width"] = Width;
                                }
                                if (Impost)
                                    NewRow["Count"] = Count * 2;
                                else
                                    NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                if (Impost)
                                    rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                                else
                                    rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void AllInsetsMarsel4(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int HeightMargin1, int InsetHeightMargin, int InsetHeightBoxMargin, int InsetWidthMargin, int InsetHeightMin, int InsetWidthMin, bool NeedSwap, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            if ((Convert.ToInt32(DT4.Rows[y]["Height"]) - 1) <= HeightMargin)
                            {
                                InsetHeightMargin = InsetHeightBoxMargin;
                            }

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]);
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]);
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            string InsetColor = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            int GroupID = Convert.ToInt32(InsetTypesDataTable.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]))[0]["GroupID"]);
                            //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");

                            if (Height > HeightMargin1)
                            {
                                Height = Height - InsetHeightMargin;
                            }
                            else
                            {
                                if (Height <= HeightMargin + 1)
                                    Height = HeightMargin - InsetHeightMargin;
                                if (Height > HeightMargin + 1 && Height <= HeightMargin1)
                                    Height = HeightMargin1 - InsetHeightMargin;
                            }
                            if (Width > HeightMargin1)
                            {
                                Width = Width - InsetWidthMargin;
                            }
                            else
                            {
                                if (Width <= HeightMargin + 1)
                                    Width = HeightMargin - InsetWidthMargin;
                                if (Width > HeightMargin + 1 && Width <= HeightMargin1)
                                    Width = HeightMargin1 - InsetWidthMargin;
                            }

                            if (GroupID == 7 || GroupID == 8)
                            {
                                //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");

                                if (Height > 100 && Width > 100)
                                    continue;
                            }
                            //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");

                            if (Width > 910)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            //if (Height <= InsetHeightMin)
                            //    Height = InsetHeightMin;
                            if (Width <= InsetWidthMin)
                                Width = InsetWidthMin;

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                if (NeedSwap)
                                {
                                    NewRow["Width"] = Height;
                                    NewRow["Height"] = Width;
                                }
                                else
                                {
                                    NewRow["Height"] = Height;
                                    NewRow["Width"] = Width;
                                }
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void AllInsetsMarsel4Impost(DataTable SourceDT, ref DataTable DestinationDT,
            int InsetHeightMargin, int InsetWidthMargin, int InsetHeightMin, int InsetWidthMin, bool NeedSwap, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID<>-1";
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DV.RowFilter = "TechnoColorID<>-1";
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DV.RowFilter = "TechnoColorID<>-1";
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DV.RowFilter = "TechnoColorID<>-1";
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) / 2 - InsetHeightMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - InsetWidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            string InsetColor = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            int GroupID = Convert.ToInt32(InsetTypesDataTable.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]))[0]["GroupID"]);
                            //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");
                            if (GroupID == 7 || GroupID == 8)
                            {
                                //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");

                                if (Height > 100 && Width > 100)
                                    continue;
                            }
                            //InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");
                            //if (Impost)
                            //{
                            //    InsetColor = InsetColor.Insert(InsetColor.Length, " 4 мм");
                            //}

                            if (Width > 910)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= InsetHeightMin)
                                Height = InsetHeightMin;
                            if (Width <= InsetWidthMin)
                                Width = InsetWidthMin;

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                if (NeedSwap)
                                {
                                    NewRow["Width"] = Height;
                                    NewRow["Height"] = Width;
                                }
                                else
                                {
                                    NewRow["Height"] = Height;
                                    NewRow["Width"] = Width;
                                }
                                NewRow["Count"] = Count * 2;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            }
                        }
                    }
                }
            }
        }

        private void AllInsetsToExcelSingly(ref HSSFWorkbook hssfworkbook,
                                            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int TotalAmount = 0;
            int AllTotalAmount = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalAmount = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void ArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (ArchDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(ArchDecorOrdersDT, true);
            DataTable DT = ArchDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = ArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Арки Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = ArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void ArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
                    HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void Assembly1ToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
           HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string PageName, ref int RowIndex, bool NeedHeader, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            if (NeedHeader)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
                cell.CellStyle = Calibri11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
                cell.CellStyle = Calibri11CS;
                if (IsPR1)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                    cell.CellStyle = CalibriBold15CS;
                }
                if (IsPR3)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                    cell.CellStyle = CalibriBold15CS;
                }
                if (IsPRU8)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                    cell.CellStyle = CalibriBold15CS;
                }

                if (DispatchDate.Length > 0)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                    cell.CellStyle = CalibriBold11CS;
                }

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
                cell.CellStyle = Calibri11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Цвет наполнителя-2");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Зачистка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "Запил витрин");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 9, "Обклад витрин");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int FType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
            {
                FType = Convert.ToInt32(DT.Rows[0]["FrontType"]);
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "Square")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        //cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        //cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        //cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    //cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    //cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    //cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    //cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    //cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    //cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void Assembly2ToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string PageName, ref int RowIndex, bool NeedHeader, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            if (NeedHeader)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
                cell.CellStyle = Calibri11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
                cell.CellStyle = Calibri11CS;
                if (IsPR1)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                    cell.CellStyle = CalibriBold15CS;
                }
                if (IsPR3)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                    cell.CellStyle = CalibriBold15CS;
                }
                if (IsPRU8)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                    cell.CellStyle = CalibriBold15CS;
                }

                if (DispatchDate.Length > 0)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                    cell.CellStyle = CalibriBold11CS;
                }

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
                cell.CellStyle = Calibri11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Цвет наполнителя-2");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Сборка");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int FType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
            {
                FType = Convert.ToInt32(DT.Rows[0]["FrontType"]);
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "Square")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        //cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        //cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        //cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    //cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    //cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    //cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    //cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    //cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    //cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void AssemblyBagetWithAngleCollect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "Length", "Height", "Width", "LeftAngle", "RightAngle", "Notes" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int Count = 0;
                string filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                    " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND LeftAngle=" + Convert.ToInt32(DT1.Rows[i]["LeftAngle"]) +
                    " AND RightAngle=" + Convert.ToInt32(DT1.Rows[i]["RightAngle"]) +
                    " AND (Notes='' OR Notes IS NULL)";
                if (DT1.Rows[i]["Notes"] != DBNull.Value && DT1.Rows[i]["Notes"].ToString().Length > 0)
                {
                    filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                      " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                      " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND LeftAngle=" + Convert.ToInt32(DT1.Rows[i]["LeftAngle"]) +
                    " AND RightAngle=" + Convert.ToInt32(DT1.Rows[i]["RightAngle"]) +
                    " AND Notes='" + DT1.Rows[i]["Notes"] + "'";
                }
                DataRow[] rows = SourceDT.Select(filter);
                if (rows.Count() == 0)
                    continue;

                foreach (DataRow item in rows)
                    Count += Convert.ToInt32(item["Count"]);

                string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                    Color += " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["DecorID"] = Convert.ToInt32(DT1.Rows[i]["DecorID"]);
                NewRow["Name"] = GetDecorName(Convert.ToInt32(DT1.Rows[i]["DecorID"]));
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "ColorID"))
                    NewRow["Color"] = Color;
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                //    NewRow["Height"] = DT1.Rows[i]["Height"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                //    NewRow["Height"] = DT1.Rows[i]["Length"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width"))
                //    NewRow["Width"] = DT1.Rows[i]["Width"];

                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width") && Convert.ToInt32(DT1.Rows[i]["Width"]) != -1)
                    NewRow["Width"] = DT1.Rows[i]["Width"];

                NewRow["Count"] = Count;
                NewRow["LeftAngle"] = DT1.Rows[i]["LeftAngle"];
                NewRow["RightAngle"] = DT1.Rows[i]["RightAngle"];
                NewRow["Notes"] = DT1.Rows[i]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Color, Height, Width, LeftAngle, RightAngle";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void AssemblyDecorCollect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "Length", "Height", "Width", "Notes" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int Count = 0;
                string filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                    " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                if (DT1.Rows[i]["Notes"] != DBNull.Value && DT1.Rows[i]["Notes"].ToString().Length > 0)
                {
                    filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                      " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                      " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) + " AND Notes='" + DT1.Rows[i]["Notes"] + "'";
                }
                DataRow[] rows = SourceDT.Select(filter);
                if (rows.Count() == 0)
                    continue;

                foreach (DataRow item in rows)
                    Count += Convert.ToInt32(item["Count"]);

                string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                    Color += " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["DecorID"] = Convert.ToInt32(DT1.Rows[i]["DecorID"]);
                NewRow["Name"] = GetDecorName(Convert.ToInt32(DT1.Rows[i]["DecorID"]));
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "ColorID"))
                    NewRow["Color"] = Color;
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                //    NewRow["Height"] = DT1.Rows[i]["Height"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                //    NewRow["Height"] = DT1.Rows[i]["Length"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width"))
                //    NewRow["Width"] = DT1.Rows[i]["Width"];

                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width") && Convert.ToInt32(DT1.Rows[i]["Width"]) != -1)
                    NewRow["Width"] = DT1.Rows[i]["Width"];

                NewRow["Count"] = Count;
                NewRow["Notes"] = DT1.Rows[i]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Color, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void AssemblyToExcel(ref HSSFWorkbook hssfworkbook,
                            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            DataTable DistFrameColorsDT = DistFrameColorsTable(Marsel1OrdersDT, true);
            DataTable DT1 = AssemblyDT.Clone();
            DataTable DT2 = AssemblyDT.Clone();
            AssemblyDT.Clear();
            FrontType = 0;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel1SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel1VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel1GridsDT, ref DT1, FrontType, ColorType, false);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Marsel5OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel5SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel5VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel5GridsDT, ref DT1, FrontType, ColorType, false);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(PortoOrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PortoSimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PortoVitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PortoGridsDT, ref DT1, FrontType, ColorType, false);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(MonteOrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), MonteSimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), MonteVitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), MonteGridsDT, ref DT1, FrontType, ColorType, false);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Marsel3OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel3SimpleDT, ref DT1, FrontType, ColorType, true);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel3VitrinaDT, ref DT2, FrontType, ColorType, true);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel3GridsDT, ref DT1, FrontType, ColorType, true);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Techno1OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno1SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno1VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno1GridsDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyLux(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno1LuxDT, ref DT1, FrontType, ColorType);
                CollectAssemblyMega(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno1MegaDT, ref DT1, FrontType, ColorType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Techno2OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno2SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno2VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno2GridsDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyLux(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno2LuxDT, ref DT1, FrontType, ColorType);
                CollectAssemblyMega(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno2MegaDT, ref DT1, FrontType, ColorType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Techno4OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno4SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno4VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno4GridsDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyLux(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno4LuxDT, ref DT1, FrontType, ColorType);
                CollectAssemblyMega(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno4MegaDT, ref DT1, FrontType, ColorType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(pFoxOrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), pFoxSimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), pFoxVitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), pFoxGridsDT, ref DT1, FrontType, ColorType, false);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(pFlorencOrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), pFlorencSimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), pFlorencVitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), pFlorencGridsDT, ref DT1, FrontType, ColorType, false);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Techno5OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno5SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno5VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno5GridsDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyLux(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Techno5LuxDT, ref DT1, FrontType, ColorType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(PR1OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PR1SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PR1VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PR1GridsDT, ref DT1, FrontType, ColorType, false);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(PR3OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PR3SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PR3VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PR3GridsDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyLux(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PR3LuxDT, ref DT1, FrontType, ColorType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(PRU8OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PRU8SimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PRU8VitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PRU8GridsDT, ref DT1, FrontType, ColorType, false);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Marsel4OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel4SimpleDT, ref DT1, FrontType, ColorType, true);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel4VitrinaDT, ref DT2, FrontType, ColorType, true);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Marsel4GridsDT, ref DT1, FrontType, ColorType, true);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Jersy110OrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Jersy110SimpleDT, ref DT1, FrontType, ColorType, true);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Jersy110VitrinaDT, ref DT2, FrontType, ColorType, true);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Jersy110GridsDT, ref DT1, FrontType, ColorType, true);
            }

            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ShervudOrdersDT, true);
            FrontType++;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                ColorType++;
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ShervudSimpleDT, ref DT1, FrontType, ColorType, false);
                CollectAssemblyVitrina(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ShervudVitrinaDT, ref DT2, FrontType, ColorType, false);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ShervudGridsDT, ref DT1, FrontType, ColorType, false);
            }
            int RowIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Зачистка");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 22 * 256);
            sheet1.SetColumnWidth(3, 9 * 256);
            sheet1.SetColumnWidth(4, 6 * 256);
            sheet1.SetColumnWidth(5, 6 * 256);
            sheet1.SetColumnWidth(6, 6 * 256);
            sheet1.SetColumnWidth(7, 13 * 256);
            sheet1.SetColumnWidth(8, 13 * 256);
            sheet1.SetColumnWidth(9, 13 * 256);

            DataTable DT = DT1.Copy();
            DataColumn Col1 = new DataColumn();
            DataColumn Col2 = new DataColumn();
            DataColumn Col3 = new DataColumn();

            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
            Col3 = DT.Columns.Add("Col3", System.Type.GetType("System.String"));
            Col1.SetOrdinal(7);
            Col2.SetOrdinal(8);
            Col3.SetOrdinal(9);

            if (DT.Rows.Count > 0)
            {
                Assembly1ToExcelSingly(ref hssfworkbook, ref sheet1,
                   CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Зачистка", ref RowIndex, true, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            DT.Dispose();
            Col1.Dispose();
            Col2.Dispose();
            Col3.Dispose();
            DT = DT2.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
            Col3 = DT.Columns.Add("Col3", System.Type.GetType("System.String"));
            Col1.SetOrdinal(7);
            Col2.SetOrdinal(8);
            Col3.SetOrdinal(9);

            if (DT.Rows.Count > 0)
            {
                Assembly1ToExcelSingly(ref hssfworkbook, ref sheet1,
                    CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Зачистка", ref RowIndex, true, IsPR1, IsPR3, IsPRU8);
            }

            RowIndex = 0;
            HSSFSheet sheet2 = hssfworkbook.CreateSheet("Сборка");
            sheet2.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet2.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet2.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet2.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet2.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet2.SetColumnWidth(0, 20 * 256);
            sheet2.SetColumnWidth(1, 11 * 256);
            sheet2.SetColumnWidth(2, 22 * 256);
            sheet2.SetColumnWidth(3, 9 * 256);
            sheet2.SetColumnWidth(4, 6 * 256);
            sheet2.SetColumnWidth(5, 6 * 256);
            sheet2.SetColumnWidth(6, 6 * 256);
            sheet2.SetColumnWidth(7, 13 * 256);
            sheet2.SetColumnWidth(8, 13 * 256);

            DT.Dispose();
            Col1.Dispose();
            DT = DT1.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(7);

            if (DT.Rows.Count > 0)
            {
                Assembly2ToExcelSingly(ref hssfworkbook, ref sheet2,
                    CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Сборка", ref RowIndex, true, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            DT.Dispose();
            Col1.Dispose();
            DT = DT2.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(7);

            if (DT.Rows.Count > 0)
            {
                Assembly2ToExcelSingly(ref hssfworkbook, ref sheet2,
                   CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Сборка", ref RowIndex, true, IsPR1, IsPR3, IsPRU8);
            }
        }

        private void BagetWithAngleAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (BagetWithAngelOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(BagetWithAngelOrdersDT, true);
            DataTable DT = BagetWithAngelOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                sheet1.SetColumnWidth(6, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    BagetWithAngleAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = BagetWithAngelOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyBagetWithAngleCollect(DT, ref BagetWithAngleAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (BagetWithAngleAssemblyDT.Rows.Count > 0)
                    {
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Багет с запилом Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                sheet1.SetColumnWidth(6, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    BagetWithAngleAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = BagetWithAngelOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyBagetWithAngleCollect(DT, ref BagetWithAngleAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (BagetWithAngleAssemblyDT.Rows.Count > 0)
                    {
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void BagetWithAngleAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
                    HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Л. угол");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "П. угол");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }
            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void CollectAdditions(DataTable SourceDT, ref DataTable DestinationDT, int coef)
        {
            if (SourceDT.Rows.Count > 0)
                AdditionsSyngly(SourceDT, ref DestinationDT, coef);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetTypeID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetTypeID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                {
                    DestinationDT.Rows[i]["InsetType"] = string.Empty;
                    DestinationDT.Rows[i]["InsetColor"] = string.Empty;
                }
            }
        }

        private void CollectAllInsets(ref DataTable DestinationDT)
        {
            if (Marsel1SimpleDT.Rows.Count > 0)
                AllInsets(Marsel1SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel1InsetHeight), Convert.ToInt32(FrontMargins.Marsel1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel1InsetMinWidth), false, true, false);

            if (Marsel5SimpleDT.Rows.Count > 0)
                AllInsets(Marsel5SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel5InsetHeight), Convert.ToInt32(FrontMargins.Marsel5InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel5InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel5InsetMinWidth), false, true, false);

            if (PortoSimpleDT.Rows.Count > 0)
                AllInsets(PortoSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.PortoInsetHeight), Convert.ToInt32(FrontMargins.PortoInsetWidth),
                    Convert.ToInt32(FrontMinSizes.PortoInsetMinHeight), Convert.ToInt32(FrontMinSizes.PortoInsetMinWidth), false, true, false);

            if (MonteSimpleDT.Rows.Count > 0)
                AllInsets(MonteSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.MonteInsetHeight), Convert.ToInt32(FrontMargins.MonteInsetWidth),
                    Convert.ToInt32(FrontMinSizes.MonteInsetMinHeight), Convert.ToInt32(FrontMinSizes.MonteInsetMinWidth), false, true, false);

            if (Techno1SimpleDT.Rows.Count > 0)
                AllInsets(Techno1SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), false, true, false);
            if (Techno2SimpleDT.Rows.Count > 0)
                AllInsets(Techno2SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), false, true, false);
            if (Techno4SimpleDT.Rows.Count > 0)
                Techno4AllInsets(Techno4SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno4Height), Convert.ToInt32(FrontMargins.Techno4Width),
                    Convert.ToInt32(FrontMargins.Techno4NarrowInsetHeight), Convert.ToInt32(FrontMargins.Techno4NarrowInsetWidth),
                    Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), false, true);
            if (pFoxSimpleDT.Rows.Count > 0)
                AllInsets(pFoxSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.pFoxInsetHeight), Convert.ToInt32(FrontMargins.pFoxInsetWidth),
                    Convert.ToInt32(FrontMinSizes.pFoxInsetMinHeight), Convert.ToInt32(FrontMinSizes.pFoxInsetMinWidth), false, true, false);
            if (pFlorencSimpleDT.Rows.Count > 0)
                AllInsets(pFlorencSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.pFlorencInsetHeight), Convert.ToInt32(FrontMargins.pFlorencInsetWidth),
                    Convert.ToInt32(FrontMinSizes.pFlorencInsetMinHeight), Convert.ToInt32(FrontMinSizes.pFlorencInsetMinWidth), false, true, false);
            if (Techno5SimpleDT.Rows.Count > 0)
                AllInsets(Techno5SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno5InsetHeight), Convert.ToInt32(FrontMargins.Techno5InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno5InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno5InsetMinWidth), false, true, false);

            if (PR1SimpleDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = PR1SimpleDT.Copy();
                for (int i = 0; i < PR1SimpleDT.Rows.Count; i++)
                {
                    object x1 = PR1SimpleDT.Rows[i]["Height"];
                    object x2 = PR1SimpleDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                AllInsets(TempPR1OrdersDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel3InsetHeight), Convert.ToInt32(FrontMargins.Marsel3InsetWidth),
                    Convert.ToInt32(FrontMinSizes.PR1InsetMinWidth), Convert.ToInt32(FrontMinSizes.PR1InsetMinHeight), false, true, false);
            }
            if (ShervudSimpleDT.Rows.Count > 0)
            {
                AllInsets(ShervudSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.ShervudInsetHeight), Convert.ToInt32(FrontMargins.ShervudInsetWidth),
                    Convert.ToInt32(FrontMinSizes.ShervudInsetMinWidth), Convert.ToInt32(FrontMinSizes.ShervudInsetMinHeight), false, true, false);
            }

            if (PR3SimpleDT.Rows.Count > 0)
                AllInsets(PR3SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.PR3InsetHeight) + 80, Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), false, true, false);

            if (PRU8SimpleDT.Rows.Count > 0)
                AllInsets(PRU8SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno1InsetHeight) + 36, Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), false, true, false);

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            DataTable DT = Marsel3SimpleDT.Clone();
            DataRow[] rows = Marsel3SimpleDT.Select("Height<1088 AND TechnoColorID=-1");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                AllInsetsMarsel3(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel3InsetHeight), Convert.ToInt32(FrontMargins.Marsel3InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel3InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel3InsetMinWidth), true, true, false);
            DT.Clear();
            rows = Marsel3SimpleDT.Select("TechnoColorID<>-1");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                AllInsetsMarsel3(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel3InsetImpostHeight), Convert.ToInt32(FrontMargins.Marsel3InsetImpostWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel3InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel3InsetMinWidth), true, true, true);
            DT.Clear();
            rows = Marsel4SimpleDT.Select("TechnoColorID=-1");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                AllInsetsMarsel4(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.Marsel4Height), Convert.ToInt32(FrontMargins.Marsel4Height1),
                    Convert.ToInt32(FrontMargins.Marsel4InsetHeight), Convert.ToInt32(FrontMargins.Marsel4BoxInsetHeight), Convert.ToInt32(FrontMargins.Marsel4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel4InsetMinWidth), false, true);
            DT.Clear();
            rows = Marsel4SimpleDT.Select("TechnoColorID<>-1");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                AllInsetsMarsel4Impost(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel4InsetImpostHeight), Convert.ToInt32(FrontMargins.Marsel4InsetImpostWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel4InsetMinWidth), false, true);

            if (Jersy110SimpleDT.Rows.Count > 0)
            {
                AllInsetsMarsel4(Jersy110SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.Jersy110Height), Convert.ToInt32(FrontMargins.Jersy110Height1),
                    Convert.ToInt32(FrontMargins.Jersy110InsetHeight), Convert.ToInt32(FrontMargins.Jersy110BoxInsetHeight), Convert.ToInt32(FrontMargins.Jersy110InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Jersy110InsetMinHeight), Convert.ToInt32(FrontMinSizes.Jersy110InsetMinWidth), false, true);
            }
            DT.Dispose();

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectAssemblyGrids(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, int ColorType, bool Impost)
        {
            DataTable DT = new DataTable();

            using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID, "ColorID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT = DV.ToTable(true, new string[] { "ColorID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                decimal Square = 0;
                int Count = 0;
                //филенки
                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetColorID"]) + " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                string TechnoColor = GetInsetTypeName(Convert.ToInt32(rows[0]["TechnoInsetTypeID"]));
                if (Convert.ToInt32(rows[0]["TechnoInsetColorID"]) != -1)
                    TechnoColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["TechnoInsetColorID"]));

                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorType;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1);
                NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));

                if (Impost && Convert.ToInt32(rows[0]["TechnoColorID"]) != -1)
                {
                    NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1) + " ИМПОСТ";
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                }

                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"])) + " (РЕШ)";
                NewRow["TechnoColor"] = TechnoColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectAssemblyLux(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, int ColorType)
        {
            DataTable DT = new DataTable();

            using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID, "ColorID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT = DV.ToTable(true, new string[] { "ColorID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                decimal Square = 0;
                int Count = 0;
                //филенки
                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetColorID"]) + " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                string TechnoColor = GetInsetTypeName(Convert.ToInt32(rows[0]["TechnoInsetTypeID"]));
                if (Convert.ToInt32(rows[0]["TechnoInsetColorID"]) != -1)
                    TechnoColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["TechnoInsetColorID"]));
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorType;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1);
                if (Convert.ToInt32(rows[0]["TechnoColorID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                if (Convert.ToInt32(rows[0]["TechnoColorID"]) != -1 && Convert.ToInt32(rows[0]["TechnoColorID"]) != Convert.ToInt32(rows[0]["ColorID"]))
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["TechnoColor"] = TechnoColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectAssemblyMega(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, int ColorType)
        {
            DataTable DT = new DataTable();

            using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID, "ColorID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT = DV.ToTable(true, new string[] { "ColorID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                decimal Square = 0;
                int Count = 0;
                //филенки
                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetColorID"]) + " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                string TechnoColor = GetInsetTypeName(Convert.ToInt32(rows[0]["TechnoInsetTypeID"]));
                if (Convert.ToInt32(rows[0]["TechnoInsetColorID"]) != -1)
                    TechnoColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["TechnoInsetColorID"]));

                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorType;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1);
                if (Convert.ToInt32(rows[0]["TechnoColorID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                {
                    if (Convert.ToInt32(rows[0]["TechnoColorID"]) != Convert.ToInt32(rows[0]["ColorID"]))
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                }
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["TechnoColor"] = TechnoColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectAssemblySimple(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, int ColorType, bool Impost)
        {
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID, "ColorID, TechnoProfileID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "TechnoProfileID",  "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string TechnoColor = string.Empty;
                //витрины
                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND TechnoProfileID=" + Convert.ToInt32(DT2.Rows[j]["TechnoProfileID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoInsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                InsetColor = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"]));
                if (Convert.ToInt32(rows[0]["InsetColorID"]) != -1)
                    InsetColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));

                TechnoColor = GetInsetTypeName(Convert.ToInt32(rows[0]["TechnoInsetTypeID"]));
                if (Convert.ToInt32(rows[0]["TechnoInsetColorID"]) != -1)
                    TechnoColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["TechnoInsetColorID"]));

                string Name = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1);
                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorType;
                NewRow["Name"] = Name;
                NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));

                if (Convert.ToInt32(rows[0]["TechnoProfileID"]) != -1)
                {
                    NewRow["Name"] = Name + " " + ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 2);
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                }

                NewRow["InsetColor"] = InsetColor;
                NewRow["TechnoColor"] = TechnoColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectAssemblyVitrina(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, int ColorType, bool Impost)
        {
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID, "ColorID, TechnoProfileID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "TechnoProfileID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = "витрина";
                string TechnoColor = string.Empty;
                //витрины
                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND TechnoProfileID=" + Convert.ToInt32(DT2.Rows[j]["TechnoProfileID"]) + 
                     " AND TechnoColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoInsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                TechnoColor = GetInsetTypeName(Convert.ToInt32(rows[0]["TechnoInsetTypeID"]));
                if (Convert.ToInt32(rows[0]["TechnoInsetColorID"]) != -1)
                    TechnoColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["TechnoInsetColorID"]));

                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                string Name = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1);
                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorType;
                NewRow["Name"] = Name;
                NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));

                if (Convert.ToInt32(rows[0]["TechnoProfileID"]) != -1)
                {
                    NewRow["Name"] = Name + " " + ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 2);
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                }

                NewRow["InsetColor"] = InsetColor;
                NewRow["TechnoColor"] = TechnoColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectDeying(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, string AdditionalName)
        {
            DataTable DT2 = new DataTable();
            //Витрины сначала
            using (DataView DV = new DataView(SourceDT, "InsetTypeID=1 AND ColorID=" + ColorID,
                "TechnoColorID, InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "PatinaID", "TechnoColorID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string filter1 = "ColorID=" + ColorID + " AND TechnoColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + ColorID + " AND TechnoColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + AdditionalName;
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));

                if (Convert.ToInt32(rows[0]["TechnoColorID"]) != -1)
                {
                    NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1) + " ИМПОСТ";
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                }

                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "Витрина";
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                NewRow["Notes"] = rows[0]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }

            using (DataView DV = new DataView(SourceDT, "InsetTypeID<>1 AND ColorID=" + ColorID,
                "TechnoColorID, InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "PatinaID", "TechnoColorID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string filter1 = "ColorID=" + ColorID + " AND TechnoColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + ColorID + " AND TechnoColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + AdditionalName;
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));

                if (Convert.ToInt32(rows[0]["TechnoColorID"]) != -1)
                {
                    NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1) + " ИМПОСТ";
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                }

                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "Витрина";
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                NewRow["Notes"] = rows[0]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectInsetsFilenkaOnly(ref DataTable DestinationDT)
        {
            if (Marsel1SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(Marsel1SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel1InsetHeight), Convert.ToInt32(FrontMargins.Marsel1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel1InsetMinWidth), true);

            if (Marsel5SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(Marsel5SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel5InsetHeight), Convert.ToInt32(FrontMargins.Marsel5InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel5InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel5InsetMinWidth), true);

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["FrontID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectInsetsGlassOnly(ref DataTable DestinationDT)
        {
            DataTable DT = Techno4MegaDT.Clone();
            DataRow[] rows = Techno4MegaDT.Select("TechnoInsetColorID=3943 AND (Height>=" + Convert.ToInt32(FrontMargins.Techno4Height) + ")");
            DT.Clear();
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                GlassOnly(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno4Height), Convert.ToInt32(FrontMargins.Techno4Width),
                    Convert.ToInt32(FrontMargins.Techno4NarrowInsetHeight), Convert.ToInt32(FrontMargins.Techno4NarrowInsetWidth),
                    Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);

            DT.Clear();
            rows = Techno1MegaDT.Select("TechnoInsetColorID=3943 AND (Height>=" + Convert.ToInt32(FrontMargins.Techno1Height) + ")");
            DT.Clear();
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                GlassOnly(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno1Height), Convert.ToInt32(FrontMargins.Techno1Width),
                    Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), true);

            DT.Clear();
            rows = Techno2MegaDT.Select("TechnoInsetColorID=3943 AND (Height>=" + Convert.ToInt32(FrontMargins.Techno2Height) + ")");
            DT.Clear();
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                GlassOnly(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno2Height), Convert.ToInt32(FrontMargins.Techno2Width),
                    Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), true);

            DT.Dispose();

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectInsetsGridsOnly(ref DataTable DestinationDT)
        {
            if (Marsel1GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Marsel1GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel1InsetHeight), Convert.ToInt32(FrontMargins.Marsel1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel1InsetMinWidth), true);
            if (Marsel5GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Marsel5GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel5InsetHeight), Convert.ToInt32(FrontMargins.Marsel5InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel5InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel5InsetMinWidth), true);
            if (PortoGridsDT.Rows.Count > 0)
                InsetsGridsOnly(PortoGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.PortoInsetHeight), Convert.ToInt32(FrontMargins.PortoInsetWidth),
                    Convert.ToInt32(FrontMinSizes.PortoInsetMinHeight), Convert.ToInt32(FrontMinSizes.PortoInsetMinWidth), true);
            if (MonteGridsDT.Rows.Count > 0)
                InsetsGridsOnly(MonteGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.MonteInsetHeight), Convert.ToInt32(FrontMargins.MonteInsetWidth),
                    Convert.ToInt32(FrontMinSizes.MonteInsetMinHeight), Convert.ToInt32(FrontMinSizes.MonteInsetMinWidth), true);
            if (Marsel3GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Marsel3GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel3InsetHeight), Convert.ToInt32(FrontMargins.Marsel3InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel3InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel3InsetMinWidth), true);
            if (Marsel4GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Marsel4GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel4InsetHeight), Convert.ToInt32(FrontMargins.Marsel4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel4InsetMinWidth), true);
            if (Jersy110GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Jersy110GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Jersy110InsetHeight), Convert.ToInt32(FrontMargins.Jersy110InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Jersy110InsetMinHeight), Convert.ToInt32(FrontMinSizes.Jersy110InsetMinWidth), true);
            if (Techno1GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Techno1GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), true);
            if (Techno2GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Techno2GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), true);
            if (Techno4GridsDT.Rows.Count > 0)
                Techno4InsetsGridsOnly(Techno4GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno4Height), Convert.ToInt32(FrontMargins.Techno4Width),
                    Convert.ToInt32(FrontMargins.Techno4NarrowInsetHeight), Convert.ToInt32(FrontMargins.Techno4NarrowInsetWidth),
                    Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);
            if (pFoxGridsDT.Rows.Count > 0)
                InsetsGridsOnly(pFoxGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.pFoxInsetHeight), Convert.ToInt32(FrontMargins.pFoxInsetWidth),
                    Convert.ToInt32(FrontMinSizes.pFoxInsetMinHeight), Convert.ToInt32(FrontMinSizes.pFoxInsetMinWidth), true);
            if (pFlorencGridsDT.Rows.Count > 0)
                InsetsGridsOnly(pFlorencGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.pFlorencInsetHeight), Convert.ToInt32(FrontMargins.pFlorencInsetWidth),
                    Convert.ToInt32(FrontMinSizes.pFlorencInsetMinHeight), Convert.ToInt32(FrontMinSizes.pFlorencInsetMinWidth), true);
            if (ShervudGridsDT.Rows.Count > 0)
                InsetsGridsOnly(ShervudGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.ShervudInsetHeight), Convert.ToInt32(FrontMargins.ShervudInsetWidth),
                    Convert.ToInt32(FrontMinSizes.ShervudInsetMinHeight), Convert.ToInt32(FrontMinSizes.ShervudInsetMinWidth), true);
            if (Techno5GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Techno5GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno5InsetHeight), Convert.ToInt32(FrontMargins.Techno5InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno5InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno5InsetMinWidth), true);
            if (PR1GridsDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = PR1GridsDT.Copy();
                for (int i = 0; i < PR1GridsDT.Rows.Count; i++)
                {
                    object x1 = PR1GridsDT.Rows[i]["Height"];
                    object x2 = PR1GridsDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                InsetsGridsOnly(TempPR1OrdersDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel3InsetHeight), Convert.ToInt32(FrontMargins.Marsel3InsetWidth),
                    Convert.ToInt32(FrontMinSizes.PR1InsetMinWidth), Convert.ToInt32(FrontMinSizes.PR1InsetMinHeight), true);
            }
            if (PR3GridsDT.Rows.Count > 0)
                InsetsGridsOnly(PR3GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.PR3InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), true);
            if (PRU8GridsDT.Rows.Count > 0)
                InsetsGridsOnly(PRU8GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), true);

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["FrontID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectInsetsLuxOnly(ref DataTable DestinationDT)
        {
            //DataTable DT = Techno1LuxDT.Clone();
            //DataRow[] rows = Techno1LuxDT.Select("(Height>=" + Convert.ToInt32(FrontMargins.Techno1Height) + ")");
            //DT.Clear();
            //foreach (DataRow item in rows)
            //    DT.Rows.Add(item.ItemArray);
            //if (DT.Rows.Count > 0)
            //    LuxOnly(DT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
            //        Convert.ToInt32(FrontMargins.LuxInsetHeight), Convert.ToInt32(FrontMargins.LuxInsetWidth), true);
            //rows = Techno2LuxDT.Select("(Height>=" + Convert.ToInt32(FrontMargins.Techno2Height) + ")");
            //DT.Clear();
            //foreach (DataRow item in rows)
            //    DT.Rows.Add(item.ItemArray);
            //if (DT.Rows.Count > 0)
            //    LuxOnly(DT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
            //        Convert.ToInt32(FrontMargins.LuxInsetHeight), Convert.ToInt32(FrontMargins.LuxInsetWidth), true);
            //rows = Techno4LuxDT.Select("(Height>=" + Convert.ToInt32(FrontMargins.Techno4Height) + ")");
            //DT.Clear();
            //foreach (DataRow item in rows)
            //    DT.Rows.Add(item.ItemArray);
            //if (DT.Rows.Count > 0)
            //    Techno4LuxOnly(DT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno4Height), Convert.ToInt32(FrontMargins.Techno4Width),
            //        Convert.ToInt32(FrontMargins.Techno4NarrowInsetHeight), Convert.ToInt32(FrontMargins.Techno4NarrowInsetWidth),
            //        Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
            //        Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);
            //rows = Techno5LuxDT.Select("(Height>=" + Convert.ToInt32(FrontMargins.Techno1Height) + ")");
            //DT.Clear();
            //foreach (DataRow item in rows)
            //    DT.Rows.Add(item.ItemArray);
            //if (DT.Rows.Count > 0)
            //    LuxOnly(DT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno5InsetHeight), Convert.ToInt32(FrontMargins.Techno5InsetWidth),
            //        Convert.ToInt32(FrontMargins.LuxInsetHeight), Convert.ToInt32(FrontMargins.LuxInsetWidth), true);

            //rows = PR3LuxDT.Select("(Height>=" + Convert.ToInt32(FrontMargins.Techno2Height) + ")");
            //DT.Clear();
            //foreach (DataRow item in rows)
            //    DT.Rows.Add(item.ItemArray);
            //if (DT.Rows.Count > 0)
            //    LuxOnly(DT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
            //        Convert.ToInt32(FrontMargins.LuxInsetHeight), Convert.ToInt32(FrontMargins.LuxInsetWidth), true);

            LuxOnly(Techno1LuxDT, ref DestinationDT,
                Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), true);

            LuxOnly(Techno2LuxDT, ref DestinationDT,
                Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), true);

            Techno4LuxOnly(Techno4LuxDT, ref DestinationDT,
                Convert.ToInt32(FrontMargins.Techno4Height), Convert.ToInt32(FrontMargins.Techno4Width),
                Convert.ToInt32(FrontMargins.Techno4NarrowInsetHeight), Convert.ToInt32(FrontMargins.Techno4NarrowInsetWidth),
                Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);

            LuxOnly(Techno5LuxDT, ref DestinationDT,
                Convert.ToInt32(FrontMargins.Techno5InsetHeight), Convert.ToInt32(FrontMargins.Techno5InsetWidth),
                Convert.ToInt32(FrontMinSizes.Techno5InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno5InsetMinWidth), true);

            LuxOnly(PR3LuxDT, ref DestinationDT,
                Convert.ToInt32(FrontMargins.PR3InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), true);

            //if (Techno2LuxDT.Rows.Count > 0)
            //    LuxOnly(Techno2LuxDT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
            //        Convert.ToInt32(FrontMargins.LuxInsetHeight), Convert.ToInt32(FrontMargins.LuxInsetWidth), true);
            //if (Techno4LuxDT.Rows.Count > 0)
            //    LuxOnly(Techno4LuxDT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
            //        Convert.ToInt32(FrontMargins.LuxInsetHeight), Convert.ToInt32(FrontMargins.LuxInsetWidth), true);
            //if (Techno5LuxDT.Rows.Count > 0)
            //    LuxOnly(Techno5LuxDT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno5InsetHeight), Convert.ToInt32(FrontMargins.Techno5InsetWidth),
            //        Convert.ToInt32(FrontMargins.LuxInsetHeight), Convert.ToInt32(FrontMargins.LuxInsetWidth), true);

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectInsetsMegaOnly(ref DataTable DestinationDT)
        {
            if (Techno1MegaDT.Rows.Count > 0)
                MegaOnly(Techno1MegaDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), true);
            if (Techno2MegaDT.Rows.Count > 0)
                MegaOnly(Techno2MegaDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), true);

            //DataTable DT = Techno4MegaDT.Clone();
            //DataRow[] rows = Techno4MegaDT.Select("TechnoInsetColorID<>128 AND (Height>=" + (Convert.ToInt32(FrontMargins.Techno4Height) + 131) + ")");
            //DT.Clear();
            //foreach (DataRow item in rows)
            //    DT.Rows.Add(item.ItemArray);
            //if (DT.Rows.Count > 0)
            //    Techno4MegaOnly(DT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Techno4Height), Convert.ToInt32(FrontMargins.Techno4Width),
            //        Convert.ToInt32(FrontMargins.Techno4NarrowInsetHeight), Convert.ToInt32(FrontMargins.Techno4NarrowInsetWidth),
            //        Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
            //        Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);
            Techno4MegaOnly(Techno4MegaDT, ref DestinationDT,
                Convert.ToInt32(FrontMargins.Techno4Height), Convert.ToInt32(FrontMargins.Techno4Width),
                Convert.ToInt32(FrontMargins.Techno4NarrowInsetHeight), Convert.ToInt32(FrontMargins.Techno4NarrowInsetWidth),
                Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["MegaCount"]) == 0)
                    DestinationDT.Rows[i]["MegaCount"] = DBNull.Value;
                if (i == 0)
                    continue;
                if (Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectInsetsPressOnly(ref DataTable DestinationDT)
        {
            DataTable DT = Marsel3SimpleDT.Clone();
            DataRow[] rows = Marsel3SimpleDT.Select("TechnoColorID=-1");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                Marsel3PressOnly(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel3InsetHeight), Convert.ToInt32(FrontMargins.Marsel3InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel3InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel3InsetMinWidth), true, false);
            DT.Clear();
            rows = Marsel3SimpleDT.Select("TechnoColorID<>-1");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                Marsel3PressOnly(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel3InsetHeight), Convert.ToInt32(FrontMargins.Marsel3InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel3InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel3InsetMinWidth), true, true);

            //if (Marsel3SimpleDT.Rows.Count > 0)
            //    Marsel3PressOnly(Marsel3SimpleDT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.Marsel3InsetHeight), Convert.ToInt32(FrontMargins.Marsel3InsetWidth),
            //        Convert.ToInt32(FrontMinSizes.Marsel3InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel3InsetMinWidth), true);
            if (Marsel4SimpleDT.Rows.Count > 0)
                InsetsPressOnly(Marsel4SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel4InsetHeight), Convert.ToInt32(FrontMargins.Marsel4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Marsel4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Marsel4InsetMinWidth), true);
            if (Jersy110SimpleDT.Rows.Count > 0)
                InsetsPressOnly(Jersy110SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Jersy110InsetHeight), Convert.ToInt32(FrontMargins.Jersy110InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Jersy110InsetMinHeight), Convert.ToInt32(FrontMinSizes.Jersy110InsetMinWidth), true);
            if (Techno1SimpleDT.Rows.Count > 0)
                InsetsPressOnly(Techno1SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), true);
            if (Techno2SimpleDT.Rows.Count > 0)
                InsetsPressOnly(Techno2SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno2InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), true);
            if (Techno4SimpleDT.Rows.Count > 0)
                InsetsPressOnly(Techno4SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);
            if (pFoxSimpleDT.Rows.Count > 0)
                InsetsPressOnly(pFoxSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);
            if (pFlorencSimpleDT.Rows.Count > 0)
                InsetsPressOnly(pFlorencSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);
            if (Techno5SimpleDT.Rows.Count > 0)
                InsetsPressOnly(Techno5SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno5InsetHeight), Convert.ToInt32(FrontMargins.Techno5InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno5InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno5InsetMinWidth), true);
            if (PR1SimpleDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = PR1SimpleDT.Copy();
                for (int i = 0; i < PR1SimpleDT.Rows.Count; i++)
                {
                    object x1 = PR1SimpleDT.Rows[i]["Height"];
                    object x2 = PR1SimpleDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                InsetsPressOnly(TempPR1OrdersDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Marsel3InsetHeight), Convert.ToInt32(FrontMargins.Marsel3InsetWidth),
                    Convert.ToInt32(FrontMinSizes.PR1InsetMinWidth), Convert.ToInt32(FrontMinSizes.PR1InsetMinHeight), true);
            }
            if (ShervudSimpleDT.Rows.Count > 0)
            {
                InsetsPressOnly(ShervudSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.ShervudInsetHeight), Convert.ToInt32(FrontMargins.ShervudInsetWidth),
                    Convert.ToInt32(FrontMinSizes.ShervudInsetMinWidth), Convert.ToInt32(FrontMinSizes.ShervudInsetMinHeight), true);
            }
            if (PR3SimpleDT.Rows.Count > 0)
                InsetsPressOnly(PR3SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.PR3InsetHeight), Convert.ToInt32(FrontMargins.Techno2InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno2InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno2InsetMinWidth), true);
            if (PRU8SimpleDT.Rows.Count > 0)
                InsetsPressOnly(PRU8SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno1InsetHeight), Convert.ToInt32(FrontMargins.Techno1InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno1InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno1InsetMinWidth), true);

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectInsetsVitrinaOnly(ref DataTable DestinationDT)
        {
            DataTable DT = Techno4MegaDT.Clone();
            DataRow[] rows = Techno4MegaDT.Select("TechnoInsetColorID=128 AND (Height>=" + Convert.ToInt32(FrontMargins.Techno4Height) + ")");
            DT.Clear();
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                Techno4MegaOnly(DT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Techno4Height), Convert.ToInt32(FrontMargins.Techno4Width),
                    Convert.ToInt32(FrontMargins.Techno4NarrowInsetHeight), Convert.ToInt32(FrontMargins.Techno4NarrowInsetWidth),
                    Convert.ToInt32(FrontMargins.Techno4InsetHeight), Convert.ToInt32(FrontMargins.Techno4InsetWidth),
                    Convert.ToInt32(FrontMinSizes.Techno4InsetMinHeight), Convert.ToInt32(FrontMinSizes.Techno4InsetMinWidth), true);
            DT.Dispose();

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectMainOrders(DataTable SourceDT, ref DataTable DestinationDT, bool Impost)
        {
            DataTable DT = new DataTable();

            using (DataView DV = new DataView(SourceDT, string.Empty, "FrontID, ColorID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT = DV.ToTable(true, new string[] { "FrontID", "ColorID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                decimal Square = 0;
                int Count = 0;

                DataRow[] rows = SourceDT.Select("FrontID=" + Convert.ToInt32(DT.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetColorID"]) +
                    " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                string TechnoColor = GetInsetTypeName(Convert.ToInt32(rows[0]["TechnoInsetTypeID"]));
                if (Convert.ToInt32(rows[0]["TechnoInsetColorID"]) != -1)
                    TechnoColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["TechnoInsetColorID"]));

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1);
                NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));

                if (Impost && Convert.ToInt32(rows[0]["TechnoColorID"]) != -1)
                {
                    NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1) + " ИМПОСТ";
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                }

                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "витрина";
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["TechnoColor"] = TechnoColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
            DestinationDT.DefaultView.Sort = "Name, FrameColor, InsetColor, TechnoColor, Height, Width";
        }

        private void CollectOrders(DataTable DistinctSizesDT, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, string FrontName, bool Impost)
        {
            int InsetTypeID = 0;
            string ColName = string.Empty;
            string FrameColor = string.Empty;
            string InsetColor = string.Empty;

            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            for (int y = 0; y < DistinctSizesDT.Rows.Count; y++)
            {
                using (DataView DV = new DataView(SourceDT))
                {
                    DT1 = DV.ToTable(true, new string[] { "ColorID", "TechnoProfileID", "TechnoColorID" });
                }
                for (int i = 0; i < DT1.Rows.Count; i++)
                {
                    using (DataView DV = new DataView(SourceDT, 
                        "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) 
                                   + " AND TechnoProfileID=" + Convert.ToInt32(DT1.Rows[i]["TechnoProfileID"]) 
                                   + " AND TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]), string.Empty, DataViewRowState.CurrentRows))
                    {
                        DT2 = DV.ToTable(true, new string[] { "InsetTypeID" });
                    }
                    for (int j = 0; j < DT2.Rows.Count; j++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]), string.Empty, DataViewRowState.CurrentRows))
                        {
                            DT3 = DV.ToTable(true, new string[] { "InsetColorID" });
                        }
                        for (int x = 0; x < DT3.Rows.Count; x++)
                        {
                            using (DataView DV = new DataView(SourceDT, "InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]), string.Empty, DataViewRowState.CurrentRows))
                            {
                                DT4 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                            }
                            for (int c = 0; c < DT4.Rows.Count; c++)
                            {
                                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                                    " AND TechnoProfileID=" + Convert.ToInt32(DT1.Rows[i]["TechnoProfileID"]) +
                                    " AND TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                                    " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) + " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                    " AND TechnoInsetTypeID=" + Convert.ToInt32(DT4.Rows[c]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT4.Rows[c]["TechnoInsetColorID"]));

                                if (rows.Count() > 0)
                                {
                                    InsetTypeID = Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]);
                                    //if (Convert.ToInt32(rows[0]["ColorID"]) == Convert.ToInt32(rows[0]["TechnoColorID"]))
                                    //    FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                                    //else
                                    //    FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));

                                    FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                                    if (Convert.ToInt32(rows[0]["TechnoColorID"]) == -1)
                                        FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                                    if (Impost && Convert.ToInt32(rows[0]["TechnoColorID"]) != -1)
                                        FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));

                                    int GroupID = Convert.ToInt32(InsetTypesDataTable.Select("InsetTypeID=" + InsetTypeID)[0]["GroupID"]);
                                    int TechnoGroupID = Convert.ToInt32(InsetTypesDataTable.Select("InsetTypeID=" + Convert.ToInt32(DT4.Rows[c]["TechnoInsetTypeID"]))[0]["GroupID"]);
                                    switch (GroupID)
                                    {
                                        case -1:
                                            InsetColor = "Витрина";
                                            break;

                                        case 7:
                                            InsetColor = "фил " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 8:
                                            InsetColor = "фил " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 3:
                                            InsetColor = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 4:
                                            InsetColor = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 16:
                                            InsetColor = "реш " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 17:
                                            InsetColor = "реш " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 18:
                                            InsetColor = "реш " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 19:
                                            InsetColor = "реш " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 12:
                                            InsetColor = "люкс " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 13:
                                            InsetColor = "мега " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"])) + "/" + GetInsetColorName(Convert.ToInt32(DT4.Rows[c]["TechnoInsetColorID"]));
                                            if (Convert.ToInt32(DT4.Rows[c]["TechnoInsetColorID"]) == 3943)
                                                InsetColor = "мега " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"])) + " Витрина";
                                            break;

                                        default:
                                            InsetColor = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;
                                    }
                                    if (TechnoGroupID == 26)
                                        InsetColor = GetInsetTypeName(Convert.ToInt32(DT4.Rows[c]["TechnoInsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));

                                    string name = FrontName;
                                    if (Convert.ToInt32(DT1.Rows[i]["TechnoProfileID"]) != -1)
                                        name = FrontName + " " + ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 2);

                                    ColName = FrameColor + "(" + InsetColor + ")_" + FrontType;
                                    if (Impost && Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) != -1)
                                        ColName = "Импост " + ColName;
                                    if (!DestinationDT.Columns.Contains(ColName))
                                        DestinationDT.Columns.Add(new DataColumn(ColName, Type.GetType("System.String")));

                                    DestinationDT.Rows[0][ColName] = name;
                                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                                        " AND TechnoProfileID=" + Convert.ToInt32(DT1.Rows[i]["TechnoProfileID"]) +
                                        " AND TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                                        " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) + " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                        " AND TechnoInsetTypeID=" + Convert.ToInt32(DT4.Rows[c]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT4.Rows[c]["TechnoInsetColorID"]) +
                                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[y]["Width"]));
                                    if (Srows.Count() > 0)
                                    {
                                        int Count = 0;
                                        foreach (DataRow item in Srows)
                                        {
                                            Count += Convert.ToInt32(item["Count"]);
                                        }

                                        DataRow[] Drows = DestinationDT.Select("Sizes='" + DistinctSizesDT.Rows[y]["Height"].ToString() + " X " + DistinctSizesDT.Rows[y]["Width"].ToString() + "'");
                                        if (Drows.Count() == 0)
                                        {
                                            DataRow NewRow = DestinationDT.NewRow();
                                            NewRow["Sizes"] = DistinctSizesDT.Rows[y]["Height"].ToString() + " X " + DistinctSizesDT.Rows[y]["Width"].ToString();
                                            NewRow["Height"] = DistinctSizesDT.Rows[y]["Height"];
                                            NewRow["Width"] = DistinctSizesDT.Rows[y]["Width"];
                                            NewRow[ColName] = Count;
                                            DestinationDT.Rows.Add(NewRow);
                                        }
                                        else
                                        {
                                            Drows[0][ColName] = Count;
                                        }
                                    }
                                }
                                else
                                    continue;
                            }
                        }
                    }
                }
            }
        }

        private void CombineComecFilenka(ref DataTable DestinationDT)
        {
            string filter = @"TechnoProfileID<>-1";

            DataTable DT = pFlorencOrdersDT.Clone();
            DataRow[] rows = pFlorencOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.pFlorencHeight) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.pFlorencInsetWidth) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.pFlorenc));
            ComecFronts(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.pFlorencHeight), 1, Front, false);

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Front, Color, ProfileType, Height DESC";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";
                if (DestinationDT.Rows[i]["BoxCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) > 0)
                {
                    if (DestinationDT.Rows[i]["Notes"] != DBNull.Value)
                    {
                        if (DestinationDT.Rows[i]["Notes"].ToString().Length > 0)
                            DestinationDT.Rows[i]["Notes"] = DestinationDT.Rows[i]["Notes"].ToString() + ", " + Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) + " шуф.";
                    }
                    else
                        DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) + " шуф.";
                }
                if (DestinationDT.Rows[i]["ImpostCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) > 0)
                {
                    if (DestinationDT.Rows[i]["Notes"] != DBNull.Value)
                    {
                        if (DestinationDT.Rows[i]["Notes"].ToString().Length > 0)
                            DestinationDT.Rows[i]["Notes"] = DestinationDT.Rows[i]["Notes"].ToString() + ", " + Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                    }
                    else
                        DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                }
                if (i == 0)
                    continue;
                if (Convert.ToInt32(DestinationDT.Rows[i]["ProfileType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ProfileType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["ColorType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["FrontType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontType"]))
                {
                    DestinationDT.Rows[i]["Front"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                }
            }
        }

        /// <summary>
        /// вкладка Stemas, нижняя таблица с импостом
        /// </summary>
        /// <param name="DestinationDT"></param>
        /// <param name="IsBox"></param>
        private void RapidTechnoInsetSimple(ref DataTable DestinationDT)
        {
            RapidTechnoInset(Marsel5OrdersDT, ref DestinationDT);
        }

        private void CombineComecSimple(ref DataTable DestinationDT)
        {
            string filter = @"TechnoProfileID<>-1";

            DataTable DT = pFlorencOrdersDT.Clone();
            DataRow[] rows = pFlorencOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.pFlorenc));
            ComecFronts(pFlorencOrdersDT, ref DestinationDT, Convert.ToInt32(FrontMargins.pFlorencHeight), 1, Front, false);
        }

        private void ComecFronts(DataTable SourceDT, ref DataTable DestinationDT, int WidthMargin, int FrontType, string Front, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistSizesTable(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }
            DT1 = OrderedTechnoFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) + " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int BoxCount = 0;
                        int VitrinaCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        int Width = Convert.ToInt32(DT2.Rows[j]["Width"]);
                        Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(Srows[0]["FrontID"]), Height, Width);

                        bool b = tuple.Item1;
                        int HeightProfile18 = tuple.Item2;
                        int Profile18 = tuple.Item3;
                        decimal HeightProfile16 = tuple.Item4;
                        int Profile16 = tuple.Item5;

                        //Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontConfigID"]), Convert.ToInt32(SourceDT.Rows[0]["TechnoProfileID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            //if (Convert.ToInt32(item["InsetTypeID"]) != 1 && (Convert.ToInt32(item["Height"]) - 10 < WidthMargin || Convert.ToInt32(item["Width"]) - 10 < WidthMargin))
                            //    BoxCount += Convert.ToInt32(item["Count"]);
                            //if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                            //    VitrinaCount += Convert.ToInt32(item["Count"]);

                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height + " AND Width=" + Width);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Width"] = Width;
                            NewRow["Count"] = Count;
                            NewRow["HeightProfile18"] = HeightProfile18;
                            NewRow["HeightProfile16"] = HeightProfile16;
                            NewRow["BoxCount"] = BoxCount * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["Profile18"] = Profile18;
                            NewRow["Profile16"] = Profile16;
                            NewRow["ImpostCount"] = 0;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            rows[0]["BoxCount"] = Convert.ToInt32(rows[0]["BoxCount"]) + BoxCount * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                        }
                    }
                }
            }
        }

        private void RapidTechnoInset(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "TechnoInsetTypeID<>-1", "ColorID, TechnoProfileID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "TechnoProfileID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                int Count = 0;
                string InsetColor = string.Empty;
                string TechnoColor = string.Empty;
                //витрины
                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND TechnoProfileID=" + Convert.ToInt32(DT2.Rows[j]["TechnoProfileID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoInsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                }

                InsetColor = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"]));
                if (Convert.ToInt32(rows[0]["InsetColorID"]) != -1)
                    InsetColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));

                TechnoColor = GetInsetTypeName(Convert.ToInt32(rows[0]["TechnoInsetTypeID"]));
                if (Convert.ToInt32(rows[0]["TechnoInsetColorID"]) != -1)
                    TechnoColor += " " + GetInsetColorName(Convert.ToInt32(rows[0]["TechnoInsetColorID"]));

                string Name = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 1);
                DataRow NewRow = DestinationDT.NewRow();
                NewRow["Name"] = Name;
                NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));

                if (Convert.ToInt32(rows[0]["TechnoProfileID"]) != -1)
                {
                    NewRow["Name"] = Name + " " + ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]), 2);
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                }

                NewRow["InsetColor"] = InsetColor;
                NewRow["TechnoColor"] = TechnoColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        public void RapidTechnoInsetToExcel(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string SheetName, string DispatchDate, string BatchName, string ClientName, ref int RowIndex)
        {
            HSSFCell cell = null;
            int DisplayIndex = 0;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, SheetName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет наполнителя-2");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;
            
            int TotalAmount = 0;

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void ComecToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string SheetName, string DispatchDate, string BatchName, string ClientName, ref int RowIndex)
        {
            HSSFCell cell = null;
            int DisplayIndex = 0;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, SheetName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Верт.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Гор.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int FType = 0;
            int PType = 0;
            int TotalAmount = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
                FType = Convert.ToInt32(DT.Rows[0]["FrontType"]);
                PType = Convert.ToInt32(DT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (DT.Rows[x]["ColorType"] != DBNull.Value)
                    CType = Convert.ToInt32(DT.Rows[x]["ColorType"]);
                if (DT.Rows[x]["FrontType"] != DBNull.Value)
                    FType = Convert.ToInt32(DT.Rows[x]["FrontType"]);
                if (DT.Rows[x]["ProfileType"] != DBNull.Value)
                    PType = Convert.ToInt32(DT.Rows[x]["ProfileType"]);
                if (x + 1 <= DT.Rows.Count - 1 &&
                    (FType != Convert.ToInt32(DT.Rows[x + 1]["FrontType"]) || PType != Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]) || CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;
                    RowIndex++;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                    PType = Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                    TotalAmount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private bool Contains(string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }

        private void Create()
        {
            DistMainOrdersDT = new DataTable();
            DistMainOrdersDT.Columns.Add(new DataColumn("MainOrderID", Type.GetType("System.Int32")));
            DistMainOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            FrontsID = new ArrayList();
            ProfileNamesDT = new DataTable();

            Marsel1VitrinaDT = new DataTable();
            Marsel1GridsDT = new DataTable();
            Marsel1SimpleDT = new DataTable();

            Marsel5VitrinaDT = new DataTable();
            Marsel5GridsDT = new DataTable();
            Marsel5SimpleDT = new DataTable();

            PortoVitrinaDT = new DataTable();
            PortoGridsDT = new DataTable();
            PortoSimpleDT = new DataTable();

            MonteVitrinaDT = new DataTable();
            MonteGridsDT = new DataTable();
            MonteSimpleDT = new DataTable();

            Marsel3VitrinaDT = new DataTable();
            Marsel3GridsDT = new DataTable();
            Marsel3SimpleDT = new DataTable();

            Marsel4VitrinaDT = new DataTable();
            Marsel4GridsDT = new DataTable();
            Marsel4SimpleDT = new DataTable();

            Jersy110VitrinaDT = new DataTable();
            Jersy110GridsDT = new DataTable();
            Jersy110SimpleDT = new DataTable();

            ShervudVitrinaDT = new DataTable();
            ShervudGridsDT = new DataTable();
            ShervudSimpleDT = new DataTable();

            Techno1VitrinaDT = new DataTable();
            Techno1GridsDT = new DataTable();
            Techno1SimpleDT = new DataTable();
            Techno1LuxDT = new DataTable();
            Techno1MegaDT = new DataTable();

            Techno2VitrinaDT = new DataTable();
            Techno2GridsDT = new DataTable();
            Techno2SimpleDT = new DataTable();
            Techno2LuxDT = new DataTable();
            Techno2MegaDT = new DataTable();

            Techno4VitrinaDT = new DataTable();
            Techno4GridsDT = new DataTable();
            Techno4SimpleDT = new DataTable();
            Techno4LuxDT = new DataTable();
            Techno4MegaDT = new DataTable();

            pFoxVitrinaDT = new DataTable();
            pFoxGridsDT = new DataTable();
            pFoxSimpleDT = new DataTable();

            pFlorencVitrinaDT = new DataTable();
            pFlorencGridsDT = new DataTable();
            pFlorencSimpleDT = new DataTable();

            Techno5VitrinaDT = new DataTable();
            Techno5GridsDT = new DataTable();
            Techno5SimpleDT = new DataTable();

            PR1VitrinaDT = new DataTable();
            PR1GridsDT = new DataTable();
            PR1SimpleDT = new DataTable();

            PR3VitrinaDT = new DataTable();
            PR3GridsDT = new DataTable();
            PR3SimpleDT = new DataTable();

            PRU8VitrinaDT = new DataTable();
            PRU8GridsDT = new DataTable();
            PRU8SimpleDT = new DataTable();

            ShervudOrdersDT = new DataTable();
            Techno1OrdersDT = new DataTable();
            Techno2OrdersDT = new DataTable();
            Techno4OrdersDT = new DataTable();
            pFoxOrdersDT = new DataTable();
            pFlorencOrdersDT = new DataTable();
            Techno5OrdersDT = new DataTable();
            PR1OrdersDT = new DataTable();
            PR3OrdersDT = new DataTable();
            PRU8OrdersDT = new DataTable();

            BagetWithAngelOrdersDT = new DataTable();
            NotArchDecorOrdersDT = new DataTable();
            ArchDecorOrdersDT = new DataTable();
            GridsDecorOrdersDT = new DataTable();

            DeyingDT = new DataTable();
            DeyingDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            DeyingDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            DecorAssemblyDT = new DataTable();
            DecorAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));

            BagetWithAngleAssemblyDT = new DataTable();
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("LeftAngle", Type.GetType("System.Decimal")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("RightAngle", Type.GetType("System.Decimal")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));

            Additions1Marsel4DT = new DataTable();
            Additions1Marsel4DT.Columns.Add(new DataColumn("InsetType", Type.GetType("System.String")));
            Additions1Marsel4DT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            Additions1Marsel4DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            Additions1Marsel4DT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            Additions1Marsel4DT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            Additions1Marsel4DT.Columns.Add(new DataColumn("TechnoInsetTypeID", Type.GetType("System.Int32")));
            Additions1Marsel4DT.Columns.Add(new DataColumn("TechnoInsetColorID", Type.GetType("System.Int32")));
            Additions2Marsel4DT = Additions1Marsel4DT.Clone();
            Additions3Marsel4DT = Additions1Marsel4DT.Clone();

            Additions1Marsel5DT = new DataTable();
            Additions1Marsel5DT.Columns.Add(new DataColumn("InsetType", Type.GetType("System.String")));
            Additions1Marsel5DT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            Additions1Marsel5DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            Additions1Marsel5DT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            Additions1Marsel5DT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            Additions1Marsel5DT.Columns.Add(new DataColumn("TechnoInsetTypeID", Type.GetType("System.Int32")));
            Additions1Marsel5DT.Columns.Add(new DataColumn("TechnoInsetColorID", Type.GetType("System.Int32")));
            Additions2Marsel5DT = Additions1Marsel5DT.Clone();
            Additions3Marsel5DT = Additions1Marsel5DT.Clone();

            TotalInfoDT = new DataTable();
            TotalInfoDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            TotalInfoDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            TotalInfoDT.Columns.Add(new DataColumn("SticksCount", Type.GetType("System.Decimal")));

            WidthMegaInsetsDT = new DataTable();
            WidthMegaInsetsDT.Columns.Add(new DataColumn("HeightMin", Type.GetType("System.Int32")));
            WidthMegaInsetsDT.Columns.Add(new DataColumn("HeightMax", Type.GetType("System.Int32")));
            WidthMegaInsetsDT.Columns.Add(new DataColumn("GlassCount", Type.GetType("System.Int32")));
            WidthMegaInsetsDT.Columns.Add(new DataColumn("MegaCount", Type.GetType("System.Int32")));

            RapidDT = new DataTable();
            RapidDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            RapidDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            RapidDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Decimal")));
            RapidDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            RapidDT.Columns.Add(new DataColumn("iCount", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("ImpostCount", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("PR1Count", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("PR2Count", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("VitrinaCount", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("ProfileType", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            ComecDT = new DataTable();
            ComecDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            ComecDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            ComecDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("HeightProfile18", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("Profile18", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("HeightProfile16", Type.GetType("System.Decimal")));
            ComecDT.Columns.Add(new DataColumn("Profile16", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            ComecDT.Columns.Add(new DataColumn("BoxCount", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("ImpostCount", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("VitrinaCount", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("ProfileType", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("FrontType", Type.GetType("System.Int32")));

            RapidTechnoInsetDT = new DataTable();
            RapidTechnoInsetDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            RapidTechnoInsetDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            RapidTechnoInsetDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            RapidTechnoInsetDT.Columns.Add(new DataColumn("TechnoColor", Type.GetType("System.String")));
            RapidTechnoInsetDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            RapidTechnoInsetDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            RapidTechnoInsetDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));

            StemasDT = new DataTable();
            StemasDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            StemasDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            StemasDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Decimal")));
            StemasDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            StemasDT.Columns.Add(new DataColumn("ImpostMargin", Type.GetType("System.Int32")));
            StemasDT.Columns.Add(new DataColumn("IsBox", Type.GetType("System.Boolean")));
            StemasDT.Columns.Add(new DataColumn("ProfileType", Type.GetType("System.Int32")));
            StemasDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            InsetDT = new DataTable();
            InsetDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            InsetDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("GlassCount", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("MegaCount", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("InsetColorID", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("TechnoInsetColorID", Type.GetType("System.Int32")));

            AssemblyDT = new DataTable();
            AssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("TechnoColor", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            AssemblyDT.Columns.Add(new DataColumn("FrontType", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            FrontsOrdersDT = new DataTable();
            FrontsOrdersDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            FrontsOrdersDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            FrontsOrdersDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            FrontsOrdersDT.Columns.Add(new DataColumn("TechnoColor", Type.GetType("System.String")));
            FrontsOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            FrontsOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            FrontsOrdersDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            FrontsOrdersDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));

            SummOrdersDT = new DataTable();
        }

        private void DeyingByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            DistMainOrdersDT.Clear();

            DistMainOrdersTable(Marsel1OrdersDT);
            DistMainOrdersTable(Marsel5OrdersDT);
            DistMainOrdersTable(PortoOrdersDT);
            DistMainOrdersTable(MonteOrdersDT);
            DistMainOrdersTable(Marsel3OrdersDT);
            DistMainOrdersTable(Marsel4OrdersDT);
            DistMainOrdersTable(Jersy110OrdersDT);
            DistMainOrdersTable(ShervudOrdersDT);
            DistMainOrdersTable(Techno1OrdersDT);
            DistMainOrdersTable(Techno2OrdersDT);
            DistMainOrdersTable(Techno4OrdersDT);
            DistMainOrdersTable(pFoxOrdersDT);
            DistMainOrdersTable(pFlorencOrdersDT);
            DistMainOrdersTable(Techno5OrdersDT);

            using (DataView DV = new DataView(DistMainOrdersDT.Copy()))
            {
                DistMainOrdersDT.Clear();
                DV.Sort = "MainOrderID ASC";
                DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }
            DataTable DT = Marsel1OrdersDT.Clone();
            DataTable DT1 = new DataTable();

            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrders.ClientID, ClientName, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }
            }

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Маркет");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 25 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);

                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    string Notes = string.Empty;

                    DeyingDT.Clear();
                    using (DataView DV = new DataView(Marsel1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Marsel1SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Marsel1GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Marsel5OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Marsel5SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Marsel5GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(PortoOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = PortoSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = PortoGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(MonteOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = MonteSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = MonteGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Marsel3OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Marsel3SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Marsel3GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Marsel4OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();

                        DataRow[] rows = Marsel4SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Marsel4VitrinaDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " Витр");

                        DT.Clear();
                        rows = Marsel4GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Jersy110OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Jersy110SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Jersy110GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ShervudOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ShervudSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ShervudGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Techno1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Techno1SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Techno1GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = Techno1LuxDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " Люкс");
                    }

                    using (DataView DV = new DataView(Techno2OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Techno2SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Techno2GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = Techno2LuxDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " Люкс");
                    }

                    using (DataView DV = new DataView(Techno4OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Techno4SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Techno4GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = Techno4LuxDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " Люкс");
                    }

                    using (DataView DV = new DataView(pFoxOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = pFoxSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = pFoxGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(pFlorencOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = pFlorencSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = pFlorencGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Techno5OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Techno5SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Techno5GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = Techno5LuxDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " Люкс");
                    }

                    string C = "Маркетинг ";
                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                        if (Convert.ToInt32(CRows[0]["ClientID"]) == 101)
                            C = "Москва-1 ";
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, WorkAssignmentID,
                            C + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, Notes, ref RowIndex);
                }
            }
        }

        private void DeyingByMainOrderToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
                                                                                                                                                                                                            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string Notes, ref int RowIndex)
        {
            DataTable TempDT = new DataTable();
            DataColumn Col1 = new DataColumn("Col1", System.Type.GetType("System.String"));
            DataColumn Col2 = new DataColumn("Col2", System.Type.GetType("System.String"));
            DataColumn Col3 = new DataColumn("Col3", System.Type.GetType("System.String"));

            if (DT.Rows.Count > 0)
            {
                TempDT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                TempDT = DT.Copy();
                Col1 = TempDT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = TempDT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                TempDT.Columns["Square"].SetOrdinal(8);

                DyeingPackingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, TempDT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Упаковка. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, ref RowIndex);
                RowIndex++;

                TempDT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                TempDT = DT.Copy();
                Col1 = TempDT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                TempDT.Columns["Square"].SetOrdinal(7);
                TempDT.Columns["Notes"].SetOrdinal(8);

                DyeingBoringToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, TempDT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Сверление. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, ref RowIndex);
            }

            RowIndex++;
        }

        private void DeyingPR1ByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            DistMainOrdersDT.Clear();

            DistMainOrdersTable(PR1OrdersDT);

            using (DataView DV = new DataView(DistMainOrdersDT.Copy()))
            {
                DistMainOrdersDT.Clear();
                DV.Sort = "MainOrderID ASC";
                DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }
            DataTable DT = PR1OrdersDT.Clone();
            DataTable DT1 = new DataTable();

            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrders.ClientID, ClientName, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }
            }

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Маркет");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 25 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);

                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    string Notes = string.Empty;

                    DeyingDT.Clear();
                    using (DataView DV = new DataView(PR1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = PR1SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = PR1GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    string C = "Маркетинг ";
                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                        if (Convert.ToInt32(CRows[0]["ClientID"]) == 101)
                            C = "Москва-1 ";
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, WorkAssignmentID,
                            C + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, Notes, ref RowIndex);
                }
            }
        }

        private void DeyingPR3ByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            DistMainOrdersDT.Clear();

            DistMainOrdersTable(PR3OrdersDT);

            using (DataView DV = new DataView(DistMainOrdersDT.Copy()))
            {
                DistMainOrdersDT.Clear();
                DV.Sort = "MainOrderID ASC";
                DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }
            DataTable DT = PR3OrdersDT.Clone();
            DataTable DT1 = new DataTable();

            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrders.ClientID, ClientName, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }
            }

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Маркет");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 25 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);

                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    string Notes = string.Empty;

                    DeyingDT.Clear();
                    using (DataView DV = new DataView(PR3OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = PR3SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = PR3GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = PR3LuxDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " Люкс");
                    }

                    string C = "Маркетинг ";
                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                        if (Convert.ToInt32(CRows[0]["ClientID"]) == 101)
                            C = "Москва-1 ";
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, WorkAssignmentID,
                            C + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, Notes, ref RowIndex);
                }
            }
        }

        private void DeyingPRU8ByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            DistMainOrdersDT.Clear();

            DistMainOrdersTable(PR3OrdersDT);

            using (DataView DV = new DataView(DistMainOrdersDT.Copy()))
            {
                DistMainOrdersDT.Clear();
                DV.Sort = "MainOrderID ASC";
                DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }
            DataTable DT = PRU8OrdersDT.Clone();
            DataTable DT1 = new DataTable();

            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrders.ClientID, ClientName, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }
            }

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Маркет");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 25 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);

                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    string Notes = string.Empty;

                    DeyingDT.Clear();
                    using (DataView DV = new DataView(PRU8OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = PRU8SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = PRU8GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    string C = "Маркетинг ";
                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                        if (Convert.ToInt32(CRows[0]["ClientID"]) == 101)
                            C = "Москва-1 ";
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, WorkAssignmentID,
                            C + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, Notes, ref RowIndex);
                }
            }
        }

        private DataTable DistFrameColorsTable(DataTable SourceDT, bool OrderASC)
        {
            int ColorID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["ColorID"].ToString(), out ColorID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["ColorID"] = ColorID;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "ColorID ASC";
                else
                    DV.Sort = "ColorID DESC";
                DT = DV.ToTable(true, new string[] { "ColorID" });
            }
            return DT;
        }

        private DataTable DistHeightTable(DataTable SourceDT, bool OrderASC)
        {
            int Height = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("ImpostMargin", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["Height"].ToString(), out Height))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["Height"] = Height;
                    NewRow["ImpostMargin"] = Row["ImpostMargin"];
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "Height ASC";
                else
                    DV.Sort = "Height DESC";
                DT = DV.ToTable(true, new string[] { "Height", "ImpostMargin" });
            }
            return DT;
        }

        private DataTable DistInsetColorsTable(DataTable SourceDT, bool OrderASC)
        {
            int InsetColorID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("InsetColorID", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                //if (Convert.ToInt32(Row["InsetTypeID"]) != 2 && Convert.ToInt32(Row["InsetTypeID"]) != 5 && Convert.ToInt32(Row["InsetTypeID"]) != 6
                //    && Convert.ToInt32(Row["InsetTypeID"]) != 9 && Convert.ToInt32(Row["InsetTypeID"]) != 10 && Convert.ToInt32(Row["InsetTypeID"]) != 11)
                //    continue;

                if (int.TryParse(Row["InsetColorID"].ToString(), out InsetColorID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["InsetColorID"] = InsetColorID;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "InsetColorID ASC";
                else
                    DV.Sort = "InsetColorID DESC";
                DT = DV.ToTable(true, new string[] { "InsetColorID" });
            }
            return DT;
        }

        private DataTable DistMainOrdersTable(DataTable SourceDT, bool OrderASC)
        {
            int MainOrderID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("MainOrderID", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "MainOrderID ASC";
                else
                    DV.Sort = "MainOrderID DESC";
                DT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }
            return DT;
        }

        private void DistMainOrdersTable(DataTable SourceDT1)
        {
            int MainOrderID = 0;
            foreach (DataRow Row in SourceDT1.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DistMainOrdersDT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DistMainOrdersDT.Rows.Add(NewRow);
                }
            }
        }

        private DataTable DistSizesTable(DataTable SourceDT, bool OrderASC)
        {
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                DataRow NewRow = DT.NewRow();
                NewRow["Height"] = Row["Height"];
                NewRow["Width"] = Row["Width"];
                DT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "Height ASC, Width ASC";
                else
                    DV.Sort = "Height DESC, Width DESC";
                DT = DV.ToTable(true, new string[] { "Height", "Width" });
            }
            return DT;
        }

        private DataTable DistWidthTable(DataTable SourceDT, bool OrderASC)
        {
            int Height = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["Width"].ToString(), out Height))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["Height"] = Height;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "Height ASC";
                else
                    DV.Sort = "Height DESC";
                DT = DV.ToTable(true, new string[] { "Height" });
            }
            return DT;
        }

        private void DyeingBoringToExcel(ref HSSFWorkbook hssfworkbook,
                                                                                            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Сверление");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void DyeingPackingToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Пленка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Упаковка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void FilenkaToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal TotalSquare = 0;
            decimal AllTotalSquare = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 2, MidpointRounding.AwayFromZero);

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void Fill()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.ProfileID",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ProfileNamesDT);
                ProfileNamesDT.Columns.Add(new DataColumn("ProfileType", Type.GetType("System.Int32")));
            }
            DecorDT = new DataTable();
            string SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID AND Enabled = 1 ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDT);
            }
            DecorParametersDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM DecorParameters",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorParametersDT);
            }

            SelectCommand = @"SELECT DISTINCT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1) ORDER BY TechStoreName";
            FrontsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }
            PatinaDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Patina",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaDataTable);
            }
            GetColorsDT();
            GetInsetColorsDT();
            InsetTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }

            StandardImpostDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM StandardImpost90",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(StandardImpostDataTable);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, InsetTypeID, PatinaID,
                ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, Notes, ImpostMargin FROM FrontsOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(Techno2OrdersDT);
                Techno2OrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));

                Marsel1OrdersDT = Techno2OrdersDT.Clone();
                Marsel1SimpleDT = Techno2OrdersDT.Clone();
                Marsel1VitrinaDT = Techno2OrdersDT.Clone();
                Marsel1GridsDT = Techno2OrdersDT.Clone();

                Marsel5OrdersDT = Techno2OrdersDT.Clone();
                Marsel5SimpleDT = Techno2OrdersDT.Clone();
                Marsel5VitrinaDT = Techno2OrdersDT.Clone();
                Marsel5GridsDT = Techno2OrdersDT.Clone();

                PortoOrdersDT = Techno2OrdersDT.Clone();
                PortoSimpleDT = Techno2OrdersDT.Clone();
                PortoVitrinaDT = Techno2OrdersDT.Clone();
                PortoGridsDT = Techno2OrdersDT.Clone();

                MonteOrdersDT = Techno2OrdersDT.Clone();
                MonteSimpleDT = Techno2OrdersDT.Clone();
                MonteVitrinaDT = Techno2OrdersDT.Clone();
                MonteGridsDT = Techno2OrdersDT.Clone();

                Marsel3OrdersDT = Techno2OrdersDT.Clone();
                Marsel3SimpleDT = Techno2OrdersDT.Clone();
                Marsel3VitrinaDT = Techno2OrdersDT.Clone();
                Marsel3GridsDT = Techno2OrdersDT.Clone();

                Marsel4OrdersDT = Techno2OrdersDT.Clone();
                Marsel4SimpleDT = Techno2OrdersDT.Clone();
                Marsel4VitrinaDT = Techno2OrdersDT.Clone();
                Marsel4GridsDT = Techno2OrdersDT.Clone();

                Jersy110OrdersDT = Techno2OrdersDT.Clone();
                Jersy110SimpleDT = Techno2OrdersDT.Clone();
                Jersy110VitrinaDT = Techno2OrdersDT.Clone();
                Jersy110GridsDT = Techno2OrdersDT.Clone();

                Techno2VitrinaDT = Techno2OrdersDT.Clone();
                Techno2GridsDT = Techno2OrdersDT.Clone();
                Techno2SimpleDT = Techno2OrdersDT.Clone();
                Techno2LuxDT = Techno2OrdersDT.Clone();
                Techno2MegaDT = Techno2OrdersDT.Clone();

                ShervudOrdersDT = Techno2OrdersDT.Clone();
                ShervudGridsDT = Techno2OrdersDT.Clone();
                ShervudSimpleDT = Techno2OrdersDT.Clone();
                ShervudVitrinaDT = Techno2OrdersDT.Clone();

                Techno1OrdersDT = Techno2OrdersDT.Clone();
                Techno1SimpleDT = Techno2OrdersDT.Clone();
                Techno1VitrinaDT = Techno2OrdersDT.Clone();
                Techno1GridsDT = Techno2OrdersDT.Clone();
                Techno1LuxDT = Techno2OrdersDT.Clone();
                Techno1MegaDT = Techno2OrdersDT.Clone();

                Techno4OrdersDT = Techno2OrdersDT.Clone();
                Techno4VitrinaDT = Techno2OrdersDT.Clone();
                Techno4SimpleDT = Techno2OrdersDT.Clone();
                Techno4GridsDT = Techno2OrdersDT.Clone();
                Techno4LuxDT = Techno2OrdersDT.Clone();

                //Techno4MegaOrdersDT = Techno2OrdersDT.Clone();
                Techno4MegaDT = Techno2OrdersDT.Clone();

                pFoxOrdersDT = Techno2OrdersDT.Clone();
                pFoxSimpleDT = Techno2OrdersDT.Clone();
                pFoxVitrinaDT = Techno2OrdersDT.Clone();
                pFoxGridsDT = Techno2OrdersDT.Clone();

                pFlorencOrdersDT = Techno2OrdersDT.Clone();
                pFlorencSimpleDT = Techno2OrdersDT.Clone();
                pFlorencVitrinaDT = Techno2OrdersDT.Clone();
                pFlorencGridsDT = Techno2OrdersDT.Clone();

                Techno5OrdersDT = Techno2OrdersDT.Clone();
                Techno5SimpleDT = Techno2OrdersDT.Clone();
                Techno5VitrinaDT = Techno2OrdersDT.Clone();
                Techno5GridsDT = Techno2OrdersDT.Clone();
                Techno5LuxDT = Techno2OrdersDT.Clone();

                PR1OrdersDT = Techno2OrdersDT.Clone();
                PR1SimpleDT = Techno2OrdersDT.Clone();
                PR1VitrinaDT = Techno2OrdersDT.Clone();
                PR1GridsDT = Techno2OrdersDT.Clone();

                PR3OrdersDT = Techno2OrdersDT.Clone();
                PR3SimpleDT = Techno2OrdersDT.Clone();
                PR3VitrinaDT = Techno2OrdersDT.Clone();
                PR3GridsDT = Techno2OrdersDT.Clone();
                PR3LuxDT = Techno2OrdersDT.Clone();

                PRU8OrdersDT = Techno2OrdersDT.Clone();
                PRU8SimpleDT = Techno2OrdersDT.Clone();
                PRU8VitrinaDT = Techno2OrdersDT.Clone();
                PRU8GridsDT = Techno2OrdersDT.Clone();
            }
            MegaInsetsNewRow(0, 170, 1, 2);
            MegaInsetsNewRow(170, 262, 2, 3);
            MegaInsetsNewRow(262, 354, 3, 4);
            MegaInsetsNewRow(354, 446, 4, 5);
            MegaInsetsNewRow(446, 538, 5, 6);
            MegaInsetsNewRow(538, 630, 6, 7);
            MegaInsetsNewRow(630, 722, 7, 8);
            MegaInsetsNewRow(722, 814, 8, 9);
            MegaInsetsNewRow(814, 906, 9, 10);
            MegaInsetsNewRow(906, 998, 10, 11);
            MegaInsetsNewRow(998, 1090, 11, 12);
            MegaInsetsNewRow(1090, 1182, 12, 13);
            MegaInsetsNewRow(1182, 1274, 13, 14);
            MegaInsetsNewRow(1274, 1366, 14, 15);
            MegaInsetsNewRow(1366, 1458, 15, 16);
            MegaInsetsNewRow(1458, 1550, 16, 17);
            MegaInsetsNewRow(1550, 1642, 17, 18);
            MegaInsetsNewRow(1642, 1734, 18, 19);
            MegaInsetsNewRow(1734, 1826, 19, 20);
            MegaInsetsNewRow(1826, 1918, 20, 21);
            MegaInsetsNewRow(1918, 2010, 21, 22);
            MegaInsetsNewRow(2010, 2102, 22, 23);
            MegaInsetsNewRow(2102, 2194, 23, 24);

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(BagetWithAngelOrdersDT);
                BagetWithAngelOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(NotArchDecorOrdersDT);
                NotArchDecorOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
                ArchDecorOrdersDT = NotArchDecorOrdersDT.Clone();
                GridsDecorOrdersDT = NotArchDecorOrdersDT.Clone();
            }
        }

        private void GetArchDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (31, 4, 18, 32) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (31, 4, 18, 32) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetBagetWithAngleOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetColorsDT()
        {
            FrameColorsDataTable = new DataTable();
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE (TechStoreGroupID = 11 OR TechStoreGroupID = 1))
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = -1;
                        NewRow["ColorName"] = "-";
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = 0;
                        NewRow["ColorName"] = "на выбор";
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["TechStoreID"]);
                        NewRow["ColorName"] = DT.Rows[i]["TechStoreName"].ToString();
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetColors WHERE GroupID IN (2,3,4,5,6,21,22)", ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow[] rows = FrameColorsDataTable.Select("ColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]));
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = FrameColorsDataTable.NewRow();
                            NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["InsetColorID"]);
                            NewRow["ColorName"] = DT.Rows[i]["InsetColorName"].ToString();
                            FrameColorsDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
        }

        private void GetCurrentDate()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT GETDATE()", ConnectionStrings.UsersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    CurrentDate = Convert.ToDateTime(DT.Rows[0][0]);
                }
            }
        }

        private string GetDecorName(int ID)
        {
            DataRow[] rows = DecorDT.Select("DecorID=" + ID);
            if (rows.Count() > 0)
                return rows[0]["Name"].ToString();
            else
                return string.Empty;
        }

        private string GetFileName(string sDestFolder, string ExcelName)
        {
            string sExtension = ".xls";
            string sFileName = ExcelName;

            int j = 1;
            while (FM.FileExist(sDestFolder + "/" + sFileName + sExtension, Configs.FTPType))
            {
                sFileName = ExcelName + "(" + j++ + ")";
            }
            sFileName = sFileName + sExtension;
            return sFileName;
        }

        private void GetFrontsOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID, Fronts Front)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, InsetTypeID, PatinaID,
                ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, Notes, ImpostMargin FROM FrontsOrders
                WHERE " +
                //(FrontConfigID IN (SELECT FrontConfigID FROM infiniu2_catalog.dbo.FrontsConfig AS F INNER JOIN
                //                         infiniu2_catalog.dbo.TechStore AS T ON F.TechnoProfileID = T.TechStoreID AND ((F.TechnoProfileID<>-1 AND SUBSTRING(T.TechStoreName, 1, 2) <> 'ПН' AND SUBSTRING(T.TechStoreName, 1, 1) <> 'Г'))) OR FrontsOrders.TechnoProfileID=-1)
                //AND
                @"FrontID=" + Convert.ToInt32(Front) +
                " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, InsetTypeID, PatinaID,
                    ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, Notes, ImpostMargin FROM FrontsOrders
                    WHERE " +
//(FrontConfigID IN (SELECT FrontConfigID FROM infiniu2_catalog.dbo.FrontsConfig AS F INNER JOIN
//                         infiniu2_catalog.dbo.TechStore AS T ON F.TechnoProfileID = T.TechStoreID AND ((F.TechnoProfileID<>-1 AND SUBSTRING(T.TechStoreName, 1, 2) <> 'ПН' AND SUBSTRING(T.TechStoreName, 1, 1) <> 'Г'))) OR FrontsOrders.TechnoProfileID=-1)
//AND
@"FrontID=" + Convert.ToInt32(Front) +
                    " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            //SelectCommand = @"SELECT FrontsOrdersID, MainOrderID, FrontID, InsetTypeID,
            //    ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, Notes, ImpostMargin FROM FrontsOrders
            //    WHERE FrontID=" + Convert.ToInt32(Front) +
            //    " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            //if (FactoryID == 2)
            //    SelectCommand = @"SELECT FrontsOrdersID, MainOrderID, FrontID, InsetTypeID,
            //        ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, Notes, ImpostMargin FROM FrontsOrders
            //        WHERE FrontID=" + Convert.ToInt32(Front) +
            //        " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
                DataRow NewRow = DestinationDT.NewRow();
                if (Convert.ToInt32(item["ImpostMargin"]) != 0)
                    bImpostMargin = true;
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetMarselOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID, Fronts Front)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, InsetTypeID, PatinaID,
                ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, Notes, ImpostMargin FROM FrontsOrders
                WHERE TechnoProfileID=-1 and FrontID=" + Convert.ToInt32(Front) +
                " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, InsetTypeID, PatinaID,
                    ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, Notes, ImpostMargin FROM FrontsOrders
                    WHERE TechnoProfileID=-1 and FrontID=" + Convert.ToInt32(Front) +
                    " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
                DataRow NewRow = DestinationDT.NewRow();
                if (Convert.ToInt32(item["ImpostMargin"]) != 0)
                    bImpostMargin = true;
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetGridFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID IN (685,686,687,688,29470,29471)");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetGridsDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetInsetColorsDT()
        {
            InsetColorsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT InsetColors.InsetColorID, InsetColors.GroupID, infiniu2_catalog.dbo.TechStore.TechStoreName AS InsetColorName FROM InsetColors" +
                " INNER JOIN infiniu2_catalog.dbo.TechStore ON InsetColors.InsetColorID = infiniu2_catalog.dbo.TechStore.TechStoreID ORDER BY TechStoreName", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetColorsDataTable);
                {
                    DataRow NewRow = InsetColorsDataTable.NewRow();
                    NewRow["InsetColorID"] = -1;
                    NewRow["GroupID"] = -1;
                    NewRow["InsetColorName"] = "-";
                    InsetColorsDataTable.Rows.Add(NewRow);
                }
                {
                    DataRow NewRow = InsetColorsDataTable.NewRow();
                    NewRow["InsetColorID"] = 0;
                    NewRow["GroupID"] = -1;
                    NewRow["InsetColorName"] = "на выбор";
                    InsetColorsDataTable.Rows.Add(NewRow);
                }
            }
        }

        private void GetLuxFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID=860");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetMainOrdersSummary(ref HSSFWorkbook hssfworkbook, HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            int MainOrderID = 0;
            int OrderNumber = 0;
            string ClientName = string.Empty;
            string DispatchDate = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;
            string SelectCommand = string.Empty;
            DataTable DistClientNamesDT = new DataTable();
            DataTable DistMainOrdersDT = new DataTable();
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.ClientID, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes AS MNotes,
                FrontsOrders.* FROM FrontsOrders
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID
                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + @")";
            //AND (FrontConfigID IN (SELECT FrontConfigID FROM infiniu2_catalog.dbo.FrontsConfig AS F INNER JOIN
            //                         infiniu2_catalog.dbo.TechStore AS T ON F.TechnoProfileID = T.TechStoreID AND((F.TechnoProfileID <> -1 AND SUBSTRING(T.TechStoreName, 1, 2) <> 'ПН' AND SUBSTRING(T.TechStoreName, 1, 1) <> 'Г'))) OR FrontsOrders.TechnoProfileID = -1)";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            if (DT.Rows.Count > 0)
            {
                DataTable TempFrontsOrdersDT = DT.Clone();
                using (DataView DV = new DataView(DT))
                {
                    DV.Sort = "ClientName";
                    DistClientNamesDT = DV.ToTable(true, new string[] { "ClientName", "ClientID" });
                }

                for (int i = 0; i < DistClientNamesDT.Rows.Count; i++)
                {
                    ClientName = DistClientNamesDT.Rows[i]["ClientName"].ToString();

                    int RowIndex = 0;
                    HSSFSheet sheet1 = hssfworkbook.CreateSheet(ClientName.Replace("/", "-"));
                    sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                    sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                    sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                    sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                    sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                    sheet1.SetColumnWidth(0, 25 * 256);
                    sheet1.SetColumnWidth(1, 11 * 256);
                    sheet1.SetColumnWidth(2, 25 * 256);
                    sheet1.SetColumnWidth(3, 15 * 256);
                    sheet1.SetColumnWidth(4, 6 * 256);
                    sheet1.SetColumnWidth(5, 6 * 256);
                    sheet1.SetColumnWidth(6, 6 * 256);

                    using (DataView DV = new DataView(DT, "ClientID=" + DistClientNamesDT.Rows[i]["ClientID"], "MainOrderID", DataViewRowState.CurrentRows))
                    {
                        DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID" });
                    }

                    for (int j = 0; j < DistMainOrdersDT.Rows.Count; j++)
                    {
                        MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[j]["MainOrderID"]);
                        DataRow[] Frows = DT.Select("MainOrderID=" + MainOrderID);
                        if (Frows.Count() == 0)
                            continue;
                        OrderNumber = Convert.ToInt32(Frows[0]["OrderNumber"]);
                        Notes = Frows[0]["MNotes"].ToString();
                        OrderName = "№" + OrderNumber.ToString() + "-" + MainOrderID;

                        TempFrontsOrdersDT.Clear();
                        FrontsOrdersDT.Clear();
                        foreach (DataRow row in Frows)
                            TempFrontsOrdersDT.Rows.Add(row.ItemArray);
                        CollectMainOrders(TempFrontsOrdersDT, ref FrontsOrdersDT, false);
                        //CollectMainOrders(TempFrontsOrdersDT, ref FrontsOrdersDT, true);

                        MainOrdersSummaryInfoToExcel(ref hssfworkbook, ref sheet1,
                               CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, FrontsOrdersDT,
                               WorkAssignmentID, DispatchDate, BatchName, ClientName, OrderName, Notes, ref RowIndex, IsPR1, IsPR3, IsPRU8);
                        RowIndex++;
                    }
                }
            }

            DistMainOrdersDT.Clear();
            DistClientNamesDT.Clear();
            DT.Clear();

            SelectCommand = @"SELECT infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.ClientID, MainOrders.DocNumber, MegaOrders.DispatchDate, MainOrders.Notes AS MNotes,
                FrontsOrdersID, TechnoProfileID, TechnoColorID, FrontsOrders.MainOrderID, FrontID, TechnoColorID, InsetTypeID,
                ColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, FrontsOrders.Notes FROM FrontsOrders
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID
                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            if (DT.Rows.Count > 0)
            {
                DataTable TempFrontsOrdersDT = DT.Clone();
                using (DataView DV = new DataView(DT))
                {
                    DV.Sort = "ClientName";
                    DistClientNamesDT = DV.ToTable(true, new string[] { "ClientName", "ClientID" });
                }

                for (int i = 0; i < DistClientNamesDT.Rows.Count; i++)
                {
                    ClientName = DistClientNamesDT.Rows[i]["ClientName"].ToString();

                    int RowIndex = 0;
                    HSSFSheet sheet1 = hssfworkbook.CreateSheet(ClientName.Replace("/", "-"));
                    sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;
                    sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                    sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                    sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                    sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                    sheet1.SetColumnWidth(0, 25 * 256);
                    sheet1.SetColumnWidth(1, 11 * 256);
                    sheet1.SetColumnWidth(2, 25 * 256);
                    sheet1.SetColumnWidth(3, 15 * 256);
                    sheet1.SetColumnWidth(4, 6 * 256);
                    sheet1.SetColumnWidth(5, 6 * 256);
                    sheet1.SetColumnWidth(6, 6 * 256);

                    using (DataView DV = new DataView(DT, "ClientID=" + DistClientNamesDT.Rows[i]["ClientID"], "MainOrderID", DataViewRowState.CurrentRows))
                    {
                        DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID" });
                    }

                    for (int j = 0; j < DistMainOrdersDT.Rows.Count; j++)
                    {
                        MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[j]["MainOrderID"]);
                        DataRow[] Frows = DT.Select("MainOrderID=" + MainOrderID);
                        if (Frows.Count() == 0)
                            continue;
                        if (Frows[0]["DispatchDate"] != DBNull.Value)
                            DispatchDate = Convert.ToDateTime(Frows[0]["DispatchDate"]).ToString("dd.MM.yyyy");
                        Notes = Frows[0]["MNotes"].ToString();
                        OrderName = Frows[0]["DocNumber"].ToString();

                        TempFrontsOrdersDT.Clear();
                        FrontsOrdersDT.Clear();
                        foreach (DataRow row in Frows)
                            TempFrontsOrdersDT.Rows.Add(row.ItemArray);
                        CollectMainOrders(TempFrontsOrdersDT, ref FrontsOrdersDT, false);
                        //CollectMainOrders(TempFrontsOrdersDT, ref FrontsOrdersDT, true);

                        MainOrdersSummaryInfoToExcel(ref hssfworkbook, ref sheet1,
                           CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, FrontsOrdersDT,
                            WorkAssignmentID, DispatchDate, BatchName, ClientName, OrderName, Notes, ref RowIndex, IsPR1, IsPR3, IsPRU8);
                        RowIndex++;
                    }
                }
            }
        }

        private string GetMarketClientName(int MainOrderID)
        {
            string ClientName = string.Empty;

            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName FROM Clients WHERE ClientID = " +
                    " (SELECT ClientID FROM infiniu2_marketingorders.dbo.MegaOrders" +
                    " WHERE MegaOrderID=(SELECT TOP 1 MegaOrderID FROM infiniu2_marketingorders.dbo.MainOrders WHERE MainOrderID = " + MainOrderID + "))",
                    ConnectionStrings.MarketingReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        ClientName = DT.Rows[0]["ClientName"].ToString();
                }
            }

            return ClientName;
        }

        private void GetMegaFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID=862");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetMegaInsetStickCount(int Height, ref int GlassCount, ref int MegaCount)
        {
            DataRow[] rows = WidthMegaInsetsDT.Select("HeightMin<=" + Height + " AND HeightMax>" + Height);
            if (rows.Count() > 0)
            {
                GlassCount = Convert.ToInt32(rows[0]["GlassCount"]);
                MegaCount = Convert.ToInt32(rows[0]["MegaCount"]);
            }
        }

        private void GetNotArchDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND NOT (ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0)) AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND NOT (ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0)) AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private string GetOrderName(int MainOrderID, int GroupType)
        {
            string name = string.Empty;
            string ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            if (GroupType == 1)
                ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            SelectCommand = @"SELECT MegaBatchID, BatchID FROM Batch WHERE BatchID IN (SELECT BatchID FROM BatchDetails WHERE MainOrderID = " + MainOrderID + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                if (DA.Fill(DT) > 0 && DT.Rows[0]["MegaBatchID"] != DBNull.Value && DT.Rows[0]["BatchID"] != DBNull.Value)
                    name = DT.Rows[0]["MegaBatchID"].ToString() + ", " + DT.Rows[0]["BatchID"] + ", " + MainOrderID;
            }
            return name;
        }

        private string GetPatinaName(int PatinaID)
        {
            string FrontType = string.Empty;
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                FrontType = Rows[0]["PatinaName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return FrontType;
        }

        private void GetProfileNames(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID, Fronts Front)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();
            {
                SelectCommand = @"SELECT infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID,1 FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.ProfileID AND infiniu2_catalog.dbo.FrontsConfig.FrontConfigID IN (SELECT FrontConfigID FROM FrontsOrders
                    WHERE FrontID=" + Convert.ToInt32(Front) +
                        " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";
                if (FactoryID == 2)
                    SelectCommand = @"SELECT infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID,1 FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.ProfileID AND infiniu2_catalog.dbo.FrontsConfig.FrontConfigID IN (SELECT FrontConfigID FROM FrontsOrders
                    WHERE FrontID=" + Convert.ToInt32(Front) +
                            " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + ")))";
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);
                }
                foreach (DataRow item in DT.Rows)
                {
                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow.ItemArray = item.ItemArray;
                    DestinationDT.Rows.Add(NewRow);
                }

                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);
                }
                foreach (DataRow item in DT.Rows)
                {
                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow.ItemArray = item.ItemArray;
                    DestinationDT.Rows.Add(NewRow);
                }
            }
            {
                SelectCommand = @"SELECT infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID,2 FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.TechnoProfileID AND infiniu2_catalog.dbo.FrontsConfig.FrontConfigID IN (SELECT FrontConfigID FROM FrontsOrders
                    WHERE FrontID=" + Convert.ToInt32(Front) +
                        " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";
                if (FactoryID == 2)
                    SelectCommand = @"SELECT infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID,2 FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.TechnoProfileID AND infiniu2_catalog.dbo.FrontsConfig.FrontConfigID IN (SELECT FrontConfigID FROM FrontsOrders
                    WHERE FrontID=" + Convert.ToInt32(Front) +
                            " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + ")))";
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);
                }
                foreach (DataRow item in DT.Rows)
                {
                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow.ItemArray = item.ItemArray;
                    DestinationDT.Rows.Add(NewRow);
                }

                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);
                }
                foreach (DataRow item in DT.Rows)
                {
                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow.ItemArray = item.ItemArray;
                    DestinationDT.Rows.Add(NewRow);
                }
            }
        }

        private void GetSimpleFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID NOT IN (1,860,862,685,686,687,688,29470,29471)");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetVitrinaFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID=1");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private string GetZOVClientName(int MainOrderID)
        {
            string ClientName = string.Empty;

            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName FROM Clients WHERE ClientID = " +
                    " (SELECT ClientID FROM infiniu2_zovorders.dbo.MainOrders WHERE MainOrderID = " + MainOrderID + ")",
                    ConnectionStrings.ZOVReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        ClientName = DT.Rows[0]["ClientName"].ToString();
                }
            }

            return ClientName;
        }

        private void GlassOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightProfilMargin, int WidthProfilMargin,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]);
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]);
                            int TempHeightMargin = HeightMargin;
                            int TempWidthMargin = WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - TempHeightMargin;
                            Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - TempWidthMargin;

                            int GlassCount = 0;
                            int MegaCount = 0;
                            string InsetColor = "стекло Лакомат";

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            foreach (DataRow item in Srows)
                            {
                                Count += Convert.ToInt32(item["Count"]);
                            }

                            GetMegaInsetStickCount(Height, ref GlassCount, ref MegaCount);

                            DataRow[] rows = DestinationDT.Select("Height=" + 30 + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = 30;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["GlassCount"] = GlassCount * Count;
                                NewRow["MegaCount"] = MegaCount * Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                                rows[0]["GlassCount"] = Convert.ToInt32(rows[0]["GlassCount"]) + GlassCount * Count;
                                rows[0]["MegaCount"] = Convert.ToInt32(rows[0]["MegaCount"]) + MegaCount * Count;
                            }
                        }
                    }
                }
            }
        }

        private void GlassOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightProfilMargin, int WidthProfilMargin, int HeightNarrowMargin, int WidthNarrowMargin,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]);
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]);
                            int TempHeightMargin = HeightMargin;
                            int TempWidthMargin = WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            if (Height < HeightProfilMargin)
                                TempHeightMargin = HeightNarrowMargin;
                            if (Width < WidthProfilMargin)
                                TempWidthMargin = WidthNarrowMargin;

                            Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - TempHeightMargin;
                            Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - TempWidthMargin;

                            int GlassCount = 0;
                            int MegaCount = 0;
                            string InsetColor = "стекло Лакомат";

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            foreach (DataRow item in Srows)
                            {
                                Count += Convert.ToInt32(item["Count"]);
                            }

                            GetMegaInsetStickCount(Height, ref GlassCount, ref MegaCount);

                            DataRow[] rows = DestinationDT.Select("Height=" + 30 + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = 30;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["GlassCount"] = GlassCount * Count;
                                NewRow["MegaCount"] = MegaCount * Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                                rows[0]["GlassCount"] = Convert.ToInt32(rows[0]["GlassCount"]) + GlassCount * Count;
                                rows[0]["MegaCount"] = Convert.ToInt32(rows[0]["MegaCount"]) + MegaCount * Count;
                            }
                        }
                    }
                }
            }
        }

        private void GridsDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (GridsDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(GridsDecorOrdersDT, true);
            DataTable DT = GridsDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Решетки1 ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = GridsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Решетки1 Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = GridsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void GridsDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
                                                                                            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void GridsToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal TotalSquare = 0;
            decimal AllTotalSquare = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        if (DT.Rows[x][y].ToString().IndexOf("3х4") != -1)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(DT.Rows[x][y].ToString());
                            cell.CellStyle = CalibriBold11CS;
                        }
                        else
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(DT.Rows[x][y].ToString());
                            cell.CellStyle = TableHeaderCS;
                        }
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 2, MidpointRounding.AwayFromZero);

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private bool HasParameter(int ProductID, string Parameter)
        {
            DataRow[] Rows = DecorParametersDT.Select("ProductID = " + ProductID);

            return Convert.ToBoolean(Rows[0][Parameter]);
        }

        private void InsetsFilenkaOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "InsetTypeID IN (2069,2070,2071,2073,2075,42066,2077,2233,3644,29043,29531,41213)", "InsetTypeID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - HeightMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);
                            string InsetColor = "фил " + GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            if (Height <= 100 || Width <= 100)
                                continue;

                            if (Height <= HeightMin || Width <= WidthMin)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void InsetsGridsOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - HeightMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            if (Height < 10 || Width < 10)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;
                            string Name = string.Empty;
                            if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 685 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 688 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 29470)
                                Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " 45 " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 686 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 687 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 29471)
                                Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " 90 " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = Name;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void InsetsPressOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - HeightMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);
                            string InsetColor = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            if (Height < 10 || Width < 10)
                                continue;

                            if (Height <= HeightMin || Width <= WidthMin)
                                continue;

                            if (Width <= 910)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void InsetToExcel(ref HSSFWorkbook hssfworkbook,
                                    HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Вставка");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 11 * 256);
            sheet1.SetColumnWidth(3, 7 * 256);
            sheet1.SetColumnWidth(4, 12 * 256);

            InsetDT.Clear();
            CollectAllInsets(ref InsetDT);

            DataTable DT = InsetDT.Copy();
            DataColumn Col1 = new DataColumn();
            DataColumn Col2 = new DataColumn();

            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            if (InsetDT.Rows.Count > 0)
            {
                AllInsetsToExcelSingly(ref hssfworkbook,
                     CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            InsetDT.Clear();
            CollectInsetsLuxOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();

            if (DT.Rows.Count > 0)
            {
                LuxToExcelSingly(ref hssfworkbook,
                  CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            InsetDT.Clear();
            CollectInsetsMegaOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();

            if (DT.Rows.Count > 0)
            {
                MegaToExcelSingly(ref hssfworkbook,
                      CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            //InsetDT.Clear();
            //CollectInsetsVitrinaOnly(ref InsetDT);

            //DT.Dispose();
            //Col1.Dispose();
            //DT = InsetDT.Copy();

            //if (DT.Rows.Count > 0)
            //{
            //    Lacomat1ToExcelSingly(ref hssfworkbook,
            //          CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex, IsPR1, IsPR3);
            //    RowIndex++;
            //    RowIndex++;
            //}

            InsetDT.Clear();
            CollectInsetsGridsOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            if (InsetDT.Rows.Count > 0)
            {
                GridsToExcelSingly(ref hssfworkbook,
                       CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            InsetDT.Clear();
            CollectInsetsGlassOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();

            if (DT.Rows.Count > 0)
            {
                Lacomat2ToExcelSingly(ref hssfworkbook,
                      CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            InsetDT.Clear();
            CollectInsetsFilenkaOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            if (DT.Rows.Count > 0)
            {
                FilenkaToExcelSingly(ref hssfworkbook,
                        CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            InsetDT.Clear();
            CollectInsetsPressOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            if (DT.Rows.Count > 0)
            {
                PressToExcelSingly(ref hssfworkbook,
                  CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            RowIndex++;
        }

        private void Lacomat1ToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "78 мм");
            cell.CellStyle = TableHeaderCS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "30 мм");
            //cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int ColumnIndex = 0;
            int TotalAmount = 0;
            int MegaCount = 0;
            int AllTotalAmount = 0;
            int AllGlassCount = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                ColumnIndex = -1;
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    MegaCount += Convert.ToInt32(DT.Rows[x]["MegaCount"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllGlassCount += Convert.ToInt32(DT.Rows[x]["MegaCount"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID"
                         || DT.Columns[y].ColumnName == "GlassCount")
                        continue;
                    ColumnIndex++;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    ColumnIndex = -1;
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID"
                                 || DT.Columns[y].ColumnName == "GlassCount")
                                continue;
                            ColumnIndex++;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        if (MegaCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                            cell.SetCellValue(MegaCount);
                            cell.CellStyle = TableHeaderCS;
                        }

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalAmount = 0;
                        MegaCount = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    ColumnIndex = -1;
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID"
                                 || DT.Columns[y].ColumnName == "GlassCount")
                                continue;
                            ColumnIndex++;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        if (MegaCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                            cell.SetCellValue(MegaCount);
                            cell.CellStyle = TableHeaderCS;
                        }
                    }
                    RowIndex++;

                    ColumnIndex = -1;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID"
                             || DT.Columns[y].ColumnName == "GlassCount")
                            continue;
                        ColumnIndex++;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    if (AllGlassCount > 0)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(AllGlassCount);
                        cell.CellStyle = TableHeaderCS;
                    }
                }
                RowIndex++;
            }
        }

        private void Lacomat2ToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "30 мм");
            //cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int ColumnIndex = 0;
            int TotalAmount = 0;
            int GlassCount = 0;
            int AllTotalAmount = 0;
            int AllGlassCount = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                ColumnIndex = -1;
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    GlassCount += Convert.ToInt32(DT.Rows[x]["GlassCount"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllGlassCount += Convert.ToInt32(DT.Rows[x]["GlassCount"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID"
                         || DT.Columns[y].ColumnName == "MegaCount" || DT.Columns[y].ColumnName == "Count")
                        continue;
                    ColumnIndex++;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    ColumnIndex = -1;
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID"
                                 || DT.Columns[y].ColumnName == "MegaCount" || DT.Columns[y].ColumnName == "Count")
                                continue;
                            ColumnIndex++;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        //cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        //cell.SetCellValue(GlassCount);
                        //cell.CellStyle = TableHeaderCS;

                        if (GlassCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                            cell.SetCellValue(GlassCount);
                            cell.CellStyle = TableHeaderCS;
                        }

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalAmount = 0;
                        GlassCount = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    ColumnIndex = -1;
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID"
                                 || DT.Columns[y].ColumnName == "MegaCount" || DT.Columns[y].ColumnName == "Count")
                                continue;
                            ColumnIndex++;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        //cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        //cell.SetCellValue(TotalAmount);
                        //cell.CellStyle = TableHeaderCS;

                        if (GlassCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                            cell.SetCellValue(GlassCount);
                            cell.CellStyle = TableHeaderCS;
                        }
                    }
                    RowIndex++;

                    ColumnIndex = -1;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID"
                             || DT.Columns[y].ColumnName == "MegaCount" || DT.Columns[y].ColumnName == "Count")
                            continue;
                        ColumnIndex++;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    //cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    //cell.SetCellValue(AllTotalAmount);
                    //cell.CellStyle = TableHeaderCS;

                    if (AllGlassCount > 0)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(AllGlassCount);
                        cell.CellStyle = TableHeaderCS;
                    }
                }
                RowIndex++;
            }
        }

        private void LuxOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;

                            if (Convert.ToInt32(DT4.Rows[y]["Height"]) <= HeightMin || Convert.ToInt32(DT4.Rows[y]["Width"]) < WidthMin)
                                continue;

                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - HeightMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);
                            int LuxCount = 0;
                            string InsetColor = "люкс " + GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            LuxCount = Convert.ToInt32(Math.Truncate(Height / 65m));
                            if (LuxCount == 0)
                                LuxCount = 1;
                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["GlassCount"] = LuxCount * Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                                rows[0]["GlassCount"] = Convert.ToInt32(rows[0]["GlassCount"]) + LuxCount * Count;
                            }
                        }
                    }
                }
            }
        }

        private void LuxToExcelSingly(ref HSSFWorkbook hssfworkbook,
                    HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "70 мм");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int TotalAmount = 0;
            int GlassCount = 0;
            int AllTotalAmount = 0;
            int AllGlassCount = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    GlassCount += Convert.ToInt32(DT.Rows[x]["GlassCount"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllGlassCount += Convert.ToInt32(DT.Rows[x]["GlassCount"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        if (GlassCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                            cell.SetCellValue(GlassCount);
                            cell.CellStyle = TableHeaderCS;
                        }

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalAmount = 0;
                        GlassCount = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        if (GlassCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                            cell.SetCellValue(GlassCount);
                            cell.CellStyle = TableHeaderCS;
                        }
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    if (AllGlassCount > 0)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(AllGlassCount);
                        cell.CellStyle = TableHeaderCS;
                    }
                }
                RowIndex++;
            }
        }

        private void MainOrdersSummaryInfoToExcel(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string OrderName, string Notes, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            RowIndex++;
            if (IsPR1)
            {
                RowIndex++;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "Заказы");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, OrderName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Цвет наполнителя-2");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Квадратура");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal TotalSquare = 0;
            int TotalAmount = 0;

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void Marsel3PressOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC, bool Impost)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height1 = Convert.ToInt32(DT4.Rows[y]["Height"]);
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - HeightMargin;

                            if (Impost)
                                Height = Convert.ToInt32(DT4.Rows[y]["Height"]) / 2 - HeightMargin;

                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);
                            string InsetColor = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            //if (Height < 10  || Width < 10)
                            //    continue;

                            if (Height <= HeightMin || Width <= WidthMin)
                                continue;

                            if (Height <= 900 && Width <= 900)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = Width;
                                NewRow["Width"] = Height;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void Martin1Fronts(ref DataTable DestinationDT)
        {
            if (Marsel1OrdersDT.Rows.Count > 0)
            {
                MartinAllFronts(Marsel1OrdersDT, ref DestinationDT, "Марсель П-016", "Марсель П-018",
                    Convert.ToInt32(FrontMargins.Marsel1Width), Convert.ToInt32(FrontMinSizes.Marsel1MinWidth), Convert.ToInt32(FrontMargins.Marsel1Height), false);
            }
            if (Marsel5OrdersDT.Rows.Count > 0)
            {
                MartinAllFronts(Marsel5OrdersDT, ref DestinationDT, "Марсель-5 П-171", "Марсель-5 П-071",
                    Convert.ToInt32(FrontMargins.Marsel5Width), Convert.ToInt32(FrontMinSizes.Marsel5MinWidth), Convert.ToInt32(FrontMargins.Marsel5Height), false);
            }
            if (PortoOrdersDT.Rows.Count > 0)
            {
                MartinAllFronts(PortoOrdersDT, ref DestinationDT, "Порто П-111", "Порто П-0111",
                    Convert.ToInt32(FrontMargins.PortoWidth), Convert.ToInt32(FrontMinSizes.PortoMinWidth), Convert.ToInt32(FrontMargins.PortoHeight), false);
            }
            if (MonteOrdersDT.Rows.Count > 0)
            {
                MartinAllFronts(MonteOrdersDT, ref DestinationDT, "Монте П-112", "Монте П-0112",
                    Convert.ToInt32(FrontMargins.MonteWidth), Convert.ToInt32(FrontMinSizes.MonteMinWidth), Convert.ToInt32(FrontMargins.MonteHeight), false);
            }
            if (Marsel3OrdersDT.Rows.Count > 0)
            {
                MartinMarsel3(Marsel3OrdersDT, ref DestinationDT, "Марсель-3 П-141", "Марсель-3 П-041",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), Convert.ToInt32(FrontMargins.Marsel3Height), false);
                MartinImpost(Marsel3OrdersDT, ref DestinationDT, "Марсель-3 П-141", "Марсель-3 П-041",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), Convert.ToInt32(FrontMargins.Marsel3Height), false);
            }
            if (Marsel4OrdersDT.Rows.Count > 0)
            {
                MartinMarsel4(Marsel4OrdersDT, ref DestinationDT, "Марсель-4 П-166", "Марсель-4 П-066",
                    Convert.ToInt32(FrontMargins.Marsel4Width), Convert.ToInt32(FrontMinSizes.Marsel4MinWidth), Convert.ToInt32(FrontMargins.Marsel4Height), Convert.ToInt32(FrontMargins.Marsel4Height1), false);
                MartinImpost(Marsel4OrdersDT, ref DestinationDT, "Марсель-4 П-166", "Марсель-4 П-066",
                    Convert.ToInt32(FrontMargins.Marsel4Width), Convert.ToInt32(FrontMinSizes.Marsel4MinWidth), Convert.ToInt32(FrontMargins.Marsel4Height), false);
            }
            if (Jersy110OrdersDT.Rows.Count > 0)
            {
                MartinMarsel4(Jersy110OrdersDT, ref DestinationDT, "Джерси П-110", "Джерси П-0110",
                    Convert.ToInt32(FrontMargins.Jersy110Width), Convert.ToInt32(FrontMinSizes.Jersy110MinWidth), Convert.ToInt32(FrontMargins.Jersy110Height), Convert.ToInt32(FrontMargins.Jersy110Height1), false);
            }
            if (Techno1OrdersDT.Rows.Count > 0)
            {
                MartinAllFronts(Techno1OrdersDT, ref DestinationDT, "Техно П-116", "Техно П-106",
                    Convert.ToInt32(FrontMargins.Techno1Width), Convert.ToInt32(FrontMinSizes.Techno1MinWidth), Convert.ToInt32(FrontMargins.Techno1Height), false);
            }
            if (Techno2OrdersDT.Rows.Count > 0)
            {
                MartinAllFronts(Techno2OrdersDT, ref DestinationDT, "Техно-2 П-216", "Техно-2 П-206",
                    Convert.ToInt32(FrontMargins.Techno2Width), Convert.ToInt32(FrontMinSizes.Techno2MinWidth), Convert.ToInt32(FrontMargins.Techno2Height), false);
            }

            if (Techno4OrdersDT.Rows.Count > 0)
                MartingTechno(Techno4OrdersDT, ref DestinationDT, "Техно-4 П-416", "Техно-4 П-406",
                    Convert.ToInt32(FrontMargins.Techno4Width), Convert.ToInt32(FrontMinSizes.Techno4MinWidth), Convert.ToInt32(FrontMargins.Techno4Height),
                    Convert.ToInt32(FrontMargins.Techno4NarrowHeight), false);
            if (pFoxOrdersDT.Rows.Count > 0)
                MartinAllFronts(pFoxOrdersDT, ref DestinationDT, "Фокс П-042-4", "Фокс П-042-4",
                    Convert.ToInt32(FrontMargins.pFoxWidth), Convert.ToInt32(FrontMinSizes.pFoxMinWidth), Convert.ToInt32(FrontMargins.pFoxHeight), false);
            if (pFlorencOrdersDT.Rows.Count > 0)
            {
                MartinFlorence1(pFlorencOrdersDT, ref DestinationDT, "П-1418-4", "П-1418-4",
                    Convert.ToInt32(FrontMargins.pFlorencWidth), Convert.ToInt32(FrontMinSizes.pFlorencMinWidth), Convert.ToInt32(FrontMargins.pFlorencHeight), false);
            }
            if (Techno5OrdersDT.Rows.Count > 0)
            {
                MartinAllFronts(Techno5OrdersDT, ref DestinationDT, "Техно-5 П-516", "Техно-5 П-506",
                    Convert.ToInt32(FrontMargins.Techno5Width), Convert.ToInt32(FrontMinSizes.Techno5MinWidth), Convert.ToInt32(FrontMargins.Techno5Height), false);
            }

            if (PR1OrdersDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = PR1OrdersDT.Copy();
                for (int i = 0; i < PR1OrdersDT.Rows.Count; i++)
                {
                    object x1 = PR1OrdersDT.Rows[i]["Height"];
                    object x2 = PR1OrdersDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                MartinPR1(TempPR1OrdersDT, ref DestinationDT, "Марсель-3 П-141", "Марсель-3 П-041",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), Convert.ToInt32(FrontMargins.Marsel3Height), false);
                MartinImpost(TempPR1OrdersDT, ref DestinationDT, "Марсель-3 П-141", "Марсель-3 П-041",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), Convert.ToInt32(FrontMargins.Marsel3Height), false);
            }
            if (ShervudOrdersDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = ShervudOrdersDT.Copy();
                for (int i = 0; i < ShervudOrdersDT.Rows.Count; i++)
                {
                    object x1 = ShervudOrdersDT.Rows[i]["Height"];
                    object x2 = ShervudOrdersDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                MartinAllFronts(TempPR1OrdersDT, ref DestinationDT, "Шервуд П-043-4", "Шервуд П-042-4",
                    Convert.ToInt32(FrontMargins.ShervudWidth), Convert.ToInt32(FrontMinSizes.ShervudMinWidth), Convert.ToInt32(FrontMargins.ShervudHeight), false);
            }
            if (PR3OrdersDT.Rows.Count > 0)
            {
                MartinPR3Rapid(PR3OrdersDT, ref DestinationDT, "Техно-2 П-216", "Техно-2 П-206",
                    Convert.ToInt32(FrontMargins.Techno2Width), Convert.ToInt32(FrontMinSizes.PR3MinWidth), Convert.ToInt32(FrontMargins.Techno2Height), false);
            }
            if (PRU8OrdersDT.Rows.Count > 0)
            {
                MartinAllFronts(PRU8OrdersDT, ref DestinationDT, "Техно-1 П-116", "Техно-1 П-106",
                    Convert.ToInt32(FrontMargins.Techno1Width), Convert.ToInt32(FrontMinSizes.Techno1MinWidth), Convert.ToInt32(FrontMargins.Techno1Height), false);
            }
            //DT.Dispose();

            string PrevName = string.Empty;
            if (DestinationDT.Rows.Count > 0)
                PrevName = DestinationDT.Rows[0]["Front"].ToString();
            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["VitrinaCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";
            }
            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["ImpostCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                if (DestinationDT.Rows[i]["iCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["iCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["iCount"]) + "-узкий проф";
                if (DestinationDT.Rows[i]["PR1Count"] != DBNull.Value && (DestinationDT.Rows[i]["PR2Count"] == DBNull.Value || Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) == 0) && Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = "ПР-1 - " + Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) + "шт";
                if ((DestinationDT.Rows[i]["PR1Count"] == DBNull.Value || Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) == 0) && DestinationDT.Rows[i]["PR2Count"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = "ПР-2 - " + Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) + "шт";
                if (DestinationDT.Rows[i]["PR1Count"] != DBNull.Value && DestinationDT.Rows[i]["PR2Count"] != DBNull.Value &&
                    Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) > 0 && Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = "ПР-1 - " + Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) + "шт/" + "ПР-2 - " + Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) + "шт";

                if (DestinationDT.Rows[i]["VitrinaCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";
                if (i == 0)
                    continue;
                string CurrentName = DestinationDT.Rows[i]["Front"].ToString();
                if (PrevName == CurrentName)
                {
                    if (Convert.ToInt32(DestinationDT.Rows[i]["ProfileType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ProfileType"]) &&
                        Convert.ToInt32(DestinationDT.Rows[i]["ColorType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorType"]))
                    {
                        PrevName = DestinationDT.Rows[i]["Front"].ToString();
                        DestinationDT.Rows[i]["Front"] = string.Empty;
                        DestinationDT.Rows[i]["Color"] = string.Empty;
                    }
                }
                else
                {
                    PrevName = DestinationDT.Rows[i]["Front"].ToString();
                }
            }
        }

        private void Martin2Fronts(ref DataTable DestinationDT)
        {
            if (pFlorencOrdersDT.Rows.Count > 0)
            {
                MartinFlorence2(pFlorencOrdersDT, ref DestinationDT, "П-1418-0", "П-1418-0",
                    Convert.ToInt32(FrontMargins.pFlorencWidth), Convert.ToInt32(FrontMinSizes.pFlorencMinWidth), Convert.ToInt32(FrontMargins.pFlorencHeight), false);
            }

            string PrevName = string.Empty;
            if (DestinationDT.Rows.Count > 0)
                PrevName = DestinationDT.Rows[0]["Front"].ToString();
            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["VitrinaCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";
            }
            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["ImpostCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                if (DestinationDT.Rows[i]["iCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["iCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["iCount"]) + "-узкий проф";
                if (DestinationDT.Rows[i]["PR1Count"] != DBNull.Value && (DestinationDT.Rows[i]["PR2Count"] == DBNull.Value || Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) == 0) && Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = "ПР-1 - " + Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) + "шт";
                if ((DestinationDT.Rows[i]["PR1Count"] == DBNull.Value || Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) == 0) && DestinationDT.Rows[i]["PR2Count"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = "ПР-2 - " + Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) + "шт";
                if (DestinationDT.Rows[i]["PR1Count"] != DBNull.Value && DestinationDT.Rows[i]["PR2Count"] != DBNull.Value &&
                    Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) > 0 && Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = "ПР-1 - " + Convert.ToInt32(DestinationDT.Rows[i]["PR1Count"]) + "шт/" + "ПР-2 - " + Convert.ToInt32(DestinationDT.Rows[i]["PR2Count"]) + "шт";

                if (DestinationDT.Rows[i]["VitrinaCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";
                if (i == 0)
                    continue;
                string CurrentName = DestinationDT.Rows[i]["Front"].ToString();
                if (PrevName == CurrentName)
                {
                    if (Convert.ToInt32(DestinationDT.Rows[i]["ProfileType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ProfileType"]) &&
                        Convert.ToInt32(DestinationDT.Rows[i]["ColorType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorType"]))
                    {
                        PrevName = DestinationDT.Rows[i]["Front"].ToString();
                        DestinationDT.Rows[i]["Front"] = string.Empty;
                        DestinationDT.Rows[i]["Color"] = string.Empty;
                    }
                }
                else
                {
                    PrevName = DestinationDT.Rows[i]["Front"].ToString();
                }
            }
        }

        private void MartinAllFronts(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                //DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int ImpostCount = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["TechnoColorID"]) != -1)
                        {
                            Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(item["FrontID"]),
                                Convert.ToInt32(item["Height"]), Convert.ToInt32(item["Width"]));
                            int HeightProfile18 = tuple.Item2;
                            int Profile18 = tuple.Item3;
                            decimal HeightProfile16 = tuple.Item4;
                            int Profile16 = tuple.Item5;

                            ImpostCount += Profile18;
                        }
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ImpostCount"] = ImpostCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                        rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + ImpostCount * 2;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int ImpostCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["TechnoColorID"]) != -1)
                        {
                            Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(item["FrontID"]),
                                Convert.ToInt32(item["Height"]), Convert.ToInt32(item["Width"]));
                            int HeightProfile18 = tuple.Item2;
                            int Profile18 = tuple.Item3;
                            decimal HeightProfile16 = tuple.Item4;
                            int Profile16 = tuple.Item5;

                            ImpostCount += Profile16;
                        }
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ImpostCount"] = ImpostCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                        rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + ImpostCount * 2;
                    }
                }
            }
        }

        private void MartinFlorence1(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "TechnoColorID=-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int ImpostCount = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["TechnoColorID"]) != -1)
                        {
                            Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(item["FrontID"]),
                                Convert.ToInt32(item["Height"]), Convert.ToInt32(item["Width"]));
                            int HeightProfile18 = tuple.Item2;
                            int Profile18 = tuple.Item3;
                            decimal HeightProfile16 = tuple.Item4;
                            int Profile16 = tuple.Item5;

                            ImpostCount += Profile18;
                        }
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ImpostCount"] = 0 * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                        rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + 0 * 2;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "TechnoColorID=-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int ImpostCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["TechnoColorID"]) != -1)
                        {
                            Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(item["FrontID"]),
                                Convert.ToInt32(item["Height"]), Convert.ToInt32(item["Width"]));
                            int HeightProfile18 = tuple.Item2;
                            int Profile18 = tuple.Item3;
                            decimal HeightProfile16 = tuple.Item4;
                            int Profile16 = tuple.Item5;

                            ImpostCount += Profile16;
                        }
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ImpostCount"] = 0 * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                        rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + 0 * 2;
                    }
                }
            }
        }

        private void MartinFlorence2(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID<>-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "TechnoColorID<>-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID<>-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int ImpostCount = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["TechnoColorID"]) != -1)
                        {
                            Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(item["FrontID"]),
                                Convert.ToInt32(item["Height"]), Convert.ToInt32(item["Width"]));
                            int HeightProfile18 = tuple.Item2;
                            int Profile18 = tuple.Item3;
                            decimal HeightProfile16 = tuple.Item4;
                            int Profile16 = tuple.Item5;

                            ImpostCount += Profile18;
                        }
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ImpostCount"] = 0 * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                        rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + 0 * 2;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "TechnoColorID<>-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID<>-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int ImpostCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["TechnoColorID"]) != -1)
                        {
                            Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(item["FrontID"]),
                                Convert.ToInt32(item["Height"]), Convert.ToInt32(item["Width"]));
                            int HeightProfile18 = tuple.Item2;
                            int Profile18 = tuple.Item3;
                            decimal HeightProfile16 = tuple.Item4;
                            int Profile16 = tuple.Item5;

                            ImpostCount += Profile16;
                        }
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ImpostCount"] = 0 * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                        rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + 0 * 2;
                    }
                }
            }
        }

        private void MartinFlorenceImpost(DataTable SourceDT, ref DataTable DestinationDT, int WidthMargin, int FrontType, string Front, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistSizesTable(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }
            DT1 = OrderedTechnoFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) + " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        int Width = Convert.ToInt32(DT2.Rows[j]["Width"]);

                        Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(Srows[0]["FrontID"]), Height, Width);

                        bool b = tuple.Item1;
                        int HeightProfile18 = tuple.Item2;
                        int Profile18 = tuple.Item3;
                        decimal HeightProfile16 = tuple.Item4;
                        int Profile16 = tuple.Item5;

                        Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontConfigID"]), 2);
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        {
                            string impostFilter = "Front='" + Front + "' AND Color='" + Color + "' AND Height=" + HeightProfile18;
                            //var filteredRows = from row in DestinationDT.AsEnumerable()
                            //                   where row.Field<decimal>("Height") == HeightProfile18 && row.Field<string>("Front") == Front
                            //                    && row.Field<string>("Color") == Color
                            //                   select row;

                            var filteredRows = DestinationDT
                                .AsEnumerable()
                                .Where(row => row.Field<decimal>("Height") == HeightProfile18
                                && row.Field<string>("Front") == Front
                                && row.Field<string>("Color") == Color);

                            DataRow[] rows = filteredRows.ToArray();
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Front"] = Front;
                                NewRow["Color"] = Color;
                                NewRow["Height"] = HeightProfile18;
                                NewRow["Count"] = Count * Profile18;
                                NewRow["ImpostCount"] = 0;
                                NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                                NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * Profile18;
                                rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + 0;
                            }
                        }

                        {
                            string impostFilter = "Front='" + Front + "' AND Color='" + Color + "' AND Height=" + HeightProfile16;

                            EnumerableRowCollection<DataRow> filteredRows = DestinationDT
                                .AsEnumerable()
                                .Where(row => row.Field<decimal>("Height") == HeightProfile16
                                && row.Field<string>("Front") == Front
                                && row.Field<string>("Color") == Color);

                            DataRow[] rows = filteredRows.ToArray();
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Front"] = Front;
                                NewRow["Color"] = Color;
                                NewRow["Height"] = HeightProfile16;
                                NewRow["Count"] = Count * Profile16;
                                NewRow["ImpostCount"] = 0;
                                NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                                NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * Profile16;
                                rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + 0;
                            }
                        }
                    }
                }
            }
        }

        private void MartingTechno(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, int HeightNarrowMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        if (Convert.ToInt32(item["Width"]) < 170)
                            iCount += Convert.ToInt32(item["Count"]);
                    }

                    //if (Height <= HeightMargin)
                    //    Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        if (Convert.ToInt32(item["Height"]) < 170)
                            iCount += Convert.ToInt32(item["Count"]);
                    }

                    if (Height < HeightMargin)
                        Height = Height - HeightNarrowMargin;
                    else
                        Height = Height - WidthMargin;
                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }
            }
        }

        private void MartinImpost(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID<>-1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    ProfileName1 = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 2) + " ИМПОСТ";
                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count;
                        NewRow["iCount"] = iCount;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount;
                    }
                }
            }
        }

        /// <summary>
        /// вкладка Мартин, нижняя таблица с импостом
        /// </summary>
        /// <param name="DestinationDT"></param>
        private void MartinImpostFronts(ref DataTable DestinationDT)
        {
            string filter = @"TechnoProfileID<>-1 ";

            DataTable DT = pFlorencOrdersDT.Clone();
            DataRow[] rows = pFlorencOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.pFlorenc));
            MartinFlorenceImpost(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.pFlorencWidth), 1, Front, false);

            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["VitrinaCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";

                if (DestinationDT.Rows[i]["ImpostCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";

                if (DestinationDT.Rows[i]["iCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["iCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["iCount"]) + "-узкий проф";
                if (i == 0)
                    continue;
                if (Convert.ToInt32(DestinationDT.Rows[i]["ProfileType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ProfileType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["ColorType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorType"]))
                {
                    DestinationDT.Rows[i]["Front"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                }
            }
        }

        private void MartinMarsel3(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                //DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int VitrinaCount = 0;
                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                            VitrinaCount += Convert.ToInt32(item["Count"]);
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["VitrinaCount"] = VitrinaCount * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int VitrinaCount = 0;
                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                            VitrinaCount += Convert.ToInt32(item["Count"]);
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["VitrinaCount"] = VitrinaCount * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }
            }
        }

        private void MartinMarsel4(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, int HeightMargin1, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                //DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height > HeightMargin1)
                    {
                    }
                    else
                    {
                        if (Height <= HeightMargin + 1)
                            Height = HeightMargin;
                        if (Height > HeightMargin + 1 && Height <= HeightMargin1)
                            Height = HeightMargin1;
                    }
                    //if (Height <= HeightMargin)
                    //    Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }
            }
        }

        private void MartinPR1(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int PR1Count = 0;
                    int PR2Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        if (Convert.ToInt32(item["FrontID"]) == Convert.ToInt32(Fronts.PR1))
                            PR1Count += Convert.ToInt32(item["Count"]);
                        if (Convert.ToInt32(item["FrontID"]) == Convert.ToInt32(Fronts.PR2))
                            PR2Count += Convert.ToInt32(item["Count"]);
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["PR1Count"] = PR1Count;
                        NewRow["PR2Count"] = PR2Count;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                        rows[0]["PR1Count"] = Convert.ToInt32(rows[0]["PR1Count"]) + PR1Count;
                        rows[0]["PR2Count"] = Convert.ToInt32(rows[0]["PR2Count"]) + PR2Count;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }
            }
        }

        private void MartinPR3Hands(DataTable SourceDT, ref DataTable DestinationDT,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (DestinationDT.Rows.Count == 0)
                        {
                            NewRow["Front"] = "Ручки ПР-3";
                        }
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void MartinPR3Rapid(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count;
                        NewRow["iCount"] = iCount;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount;
                    }
                }
            }
        }

        private void MartinPRU8(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                //DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            int ProfileType = 0;
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";
                if (!OrderASC)
                    SizesASC = "Height DESC";
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }

                ProfileType++;
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int iCount = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    string Notes = string.Empty;
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["iCount"] = iCount * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        rows[0]["iCount"] = Convert.ToInt32(rows[0]["iCount"]) + iCount * 2;
                    }
                }
            }
        }

        private void MartinPRU8Hands(DataTable SourceDT, ref DataTable DestinationDT,
            int WidthMargin, int WidthMin, int HeightMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Width ASC";
                if (!OrderASC)
                    SizesASC = "Width DESC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (DestinationDT.Rows.Count == 0)
                        {
                            NewRow["Front"] = "ПРУ-8";
                            NewRow["Color"] = "Ручки";
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                    }
                }
            }
        }

        private void MartinToExcel1(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1, ref int RowIndex, HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            if (DT.Rows.Count == 0)
                return;

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "MARTIN");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int PType = 0;
            int TotalAmount = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (RapidDT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(RapidDT.Rows[0]["ColorType"]);
                PType = Convert.ToInt32(RapidDT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < RapidDT.Rows.Count; x++)
            {
                if (RapidDT.Rows[x]["Count"] != DBNull.Value && RapidDT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(RapidDT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(RapidDT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    TotalAmount += Convert.ToInt32(RapidDT.Rows[x]["Count"]);
                    AllTotalAmount += Convert.ToInt32(RapidDT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(RapidDT.Rows[x]["Count"]);
                }

                for (int y = 0; y < RapidDT.Columns.Count; y++)
                {
                    if (RapidDT.Columns[y].ColumnName == "ImpostCount" || RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                        || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                        continue;
                    Type t = RapidDT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(RapidDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(RapidDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(RapidDT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= RapidDT.Rows.Count - 1 && (PType != Convert.ToInt32(RapidDT.Rows[x + 1]["ProfileType"]) || CType != Convert.ToInt32(RapidDT.Rows[x + 1]["ColorType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < RapidDT.Columns.Count; y++)
                    {
                        if (RapidDT.Columns[y].ColumnName == "ImpostCount" || RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                            || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;

                    CType = Convert.ToInt32(RapidDT.Rows[x + 1]["ColorType"]);
                    PType = Convert.ToInt32(RapidDT.Rows[x + 1]["ProfileType"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                    TotalAmount = 0;
                }

                if (x == RapidDT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < RapidDT.Columns.Count; y++)
                    {
                        if (RapidDT.Columns[y].ColumnName == "ImpostCount" || RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                            || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < RapidDT.Columns.Count; y++)
                    {
                        if (RapidDT.Columns[y].ColumnName == "ImpostCount" || RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                            || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void MartinToExcel2(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1, ref int RowIndex,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "француз<100");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int PType = 0;
            int TotalAmount = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (RapidDT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(RapidDT.Rows[0]["ColorType"]);
                PType = Convert.ToInt32(RapidDT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < RapidDT.Rows.Count; x++)
            {
                if (RapidDT.Rows[x]["Count"] != DBNull.Value && RapidDT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(RapidDT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(RapidDT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    TotalAmount += Convert.ToInt32(RapidDT.Rows[x]["Count"]);
                    AllTotalAmount += Convert.ToInt32(RapidDT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(RapidDT.Rows[x]["Count"]);
                }

                for (int y = 0; y < RapidDT.Columns.Count; y++)
                {
                    if (RapidDT.Columns[y].ColumnName == "ImpostCount" || RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                        || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                        continue;
                    Type t = RapidDT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(RapidDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(RapidDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(RapidDT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= RapidDT.Rows.Count - 1 && (PType != Convert.ToInt32(RapidDT.Rows[x + 1]["ProfileType"]) || CType != Convert.ToInt32(RapidDT.Rows[x + 1]["ColorType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < RapidDT.Columns.Count; y++)
                    {
                        if (RapidDT.Columns[y].ColumnName == "ImpostCount" || RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                            || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;

                    CType = Convert.ToInt32(RapidDT.Rows[x + 1]["ColorType"]);
                    PType = Convert.ToInt32(RapidDT.Rows[x + 1]["ProfileType"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                    TotalAmount = 0;
                }

                if (x == RapidDT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < RapidDT.Columns.Count; y++)
                    {
                        if (RapidDT.Columns[y].ColumnName == "ImpostCount" || RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                            || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < RapidDT.Columns.Count; y++)
                    {
                        if (RapidDT.Columns[y].ColumnName == "ImpostCount" || RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                            || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void MegaInsetsNewRow(int HeightMin, int HeightMax, int GlassCount, int MegaCount)
        {
            DataRow NewRow = WidthMegaInsetsDT.NewRow();
            NewRow["HeightMin"] = HeightMin;
            NewRow["HeightMax"] = HeightMax;
            NewRow["GlassCount"] = GlassCount;
            NewRow["MegaCount"] = MegaCount;
            WidthMegaInsetsDT.Rows.Add(NewRow);
        }

        private void MegaOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            //if (Convert.ToInt32(DT4.Rows[y]["Height"]) <= HeightMin)
                            //    continue;

                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - HeightMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            //if (Height < 20 || Width < 20)
                            //    continue;

                            int Count = 0;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);
                            int GlassCount = 0;
                            int MegaCount = 0;
                            string InsetColor = "мега " + GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            foreach (DataRow item in Srows)
                            {
                                Count += Convert.ToInt32(item["Count"]);
                            }

                            GetMegaInsetStickCount(Height, ref GlassCount, ref MegaCount);

                            if (TechnoInsetColorID == 3943)
                            {
                                GlassCount = 0;
                                InsetColor += " вит";
                            }
                            else
                                InsetColor += "/" + GetInsetColorName(TechnoInsetColorID);

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["GlassCount"] = GlassCount * Count;
                                NewRow["MegaCount"] = MegaCount * Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                                rows[0]["GlassCount"] = Convert.ToInt32(rows[0]["GlassCount"]) + GlassCount * Count;
                                rows[0]["MegaCount"] = Convert.ToInt32(rows[0]["MegaCount"]) + MegaCount * Count;
                            }
                        }
                    }
                }
            }
        }

        private void MegaToExcelSingly(ref HSSFWorkbook hssfworkbook,
                            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "30 мм");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "78 мм");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int TotalAmount = 0;
            int GlassCount = 0;
            int MegaCount = 0;
            int AllTotalAmount = 0;
            int AllGlassCount = 0;
            int AllMegaCount = 0;
            string str = string.Empty;

            int AType = -1;
            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                AType = Convert.ToInt32(DT.Rows[0]["TechnoInsetColorID"]);
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    GlassCount += Convert.ToInt32(DT.Rows[x]["GlassCount"]);
                    if (DT.Rows[x]["MegaCount"] != DBNull.Value)
                        MegaCount += Convert.ToInt32(DT.Rows[x]["MegaCount"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllGlassCount += Convert.ToInt32(DT.Rows[x]["GlassCount"]);
                    if (DT.Rows[x]["MegaCount"] != DBNull.Value)
                        AllMegaCount += Convert.ToInt32(DT.Rows[x]["MegaCount"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]) || AType != Convert.ToInt32(DT.Rows[x + 1]["TechnoInsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        if (GlassCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                            cell.SetCellValue(GlassCount);
                            cell.CellStyle = TableHeaderCS;
                        }
                        else
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            //continue;
                        }

                        if (MegaCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                            cell.SetCellValue(MegaCount);
                            cell.CellStyle = TableHeaderCS;
                        }
                        else
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            //continue;
                        }

                        AType = Convert.ToInt32(DT.Rows[x + 1]["TechnoInsetColorID"]);
                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalAmount = 0;
                        GlassCount = 0;
                        MegaCount = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        if (GlassCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                            cell.SetCellValue(GlassCount);
                            cell.CellStyle = TableHeaderCS;
                        }
                        else
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            //continue;
                        }

                        if (MegaCount > 0)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                            cell.SetCellValue(MegaCount);
                            cell.CellStyle = TableHeaderCS;
                        }
                        else
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            //continue;
                        }
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    if (AllGlassCount > 0)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(AllGlassCount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    else
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (AllMegaCount > 0)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(AllMegaCount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    else
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                RowIndex++;
            }
        }

        private void NotArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (NotArchDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(NotArchDecorOrdersDT, true);
            DataTable DT = NotArchDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = NotArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = NotArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void NotArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
                    HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }
            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private DataTable OrderedTechnoFrameColors(DataTable SourceDT)
        {
            DataTable OrderedDT = SourceDT.Copy();
            OrderedDT.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));

            for (int i = 0; i < OrderedDT.Rows.Count; i++)
                OrderedDT.Rows[i]["ColorName"] = GetColorName(Convert.ToInt32(OrderedDT.Rows[i]["TechnoColorID"]));

            using (DataView DV = new DataView(OrderedDT.Copy()))
            {
                DV.Sort = "ColorName";
                OrderedDT.Clear();
                OrderedDT = DV.ToTable();
            }

            return OrderedDT;
        }

        private void OrdersSummaryInfoToExcel(ref HSSFWorkbook hssfworkbook,
           HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, bool IsPR1, bool IsPR2, bool IsPR3, bool IsPRU8)
        {
            int RowIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Заказы");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);
            decimal AllSquare = 0;
            string FrontName = string.Empty;
            if (Marsel1OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(Marsel1OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryMarsel1Orders(Marsel1OrdersDT, Marsel1SimpleDT, Marsel1VitrinaDT, Marsel1GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                     CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            if (Marsel5OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(Marsel5OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryMarsel5Orders(Marsel5OrdersDT, new DataView(Marsel5SimpleDT, "TechnoInsetTypeID=-1", "", DataViewRowState.CurrentRows).ToTable(), Marsel5VitrinaDT, Marsel5GridsDT, new DataView(Marsel5SimpleDT, "TechnoInsetTypeID<>-1", "", DataViewRowState.CurrentRows).ToTable(), FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                     CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (PortoOrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(PortoOrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryMarsel1Orders(PortoOrdersDT, PortoSimpleDT, PortoVitrinaDT, PortoGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                     CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (MonteOrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(MonteOrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryMarsel1Orders(MonteOrdersDT, MonteSimpleDT, MonteVitrinaDT, MonteGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                     CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (Marsel3OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(Marsel3OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryMarsel3Orders(Marsel3OrdersDT, Marsel3SimpleDT, Marsel3VitrinaDT, Marsel3GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                       CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (Marsel4OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(Marsel4OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryMarsel3Orders(Marsel4OrdersDT, Marsel4SimpleDT, Marsel4VitrinaDT, Marsel4GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                       CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (Jersy110OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(Jersy110OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryMarsel3Orders(Jersy110OrdersDT, Jersy110SimpleDT, Jersy110VitrinaDT, Jersy110GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                       CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (ShervudOrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(ShervudOrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryShervudOrders(ShervudOrdersDT, ShervudSimpleDT, ShervudVitrinaDT, ShervudGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                   CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (Techno1OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(Techno1OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryTechno1Orders(Techno1OrdersDT, Techno1SimpleDT, Techno1VitrinaDT, Techno1GridsDT, Techno1LuxDT, Techno1MegaDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                   CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (Techno2OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(Techno2OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryTechno2Orders(Techno2OrdersDT, Techno2SimpleDT, Techno2VitrinaDT, Techno2GridsDT, Techno2LuxDT, Techno2MegaDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                   CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (Techno4OrdersDT.Rows.Count > 0)
            {
                if (Techno4OrdersDT.Rows.Count > 0)
                    FrontName = ProfileName(Convert.ToInt32(Techno4OrdersDT.Rows[0]["FrontConfigID"]), 1);
                //if (Techno4MegaOrdersDT.Rows.Count > 0)
                //    FrontName = ProfileName(Convert.ToInt32(Techno4MegaOrdersDT.Rows[0]["FrontConfigID"]));

                DataTable DT = Techno4OrdersDT.Copy();
                //foreach (DataRow item in Techno4MegaOrdersDT.Rows)
                //    DT.Rows.Add(item.ItemArray);
                SummaryTechno4Orders(DT, Techno4SimpleDT, Techno4VitrinaDT, Techno4GridsDT, Techno4LuxDT, Techno4MegaDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                   CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (pFoxOrdersDT.Rows.Count > 0)
            {
                if (pFoxOrdersDT.Rows.Count > 0)
                    FrontName = ProfileName(Convert.ToInt32(pFoxOrdersDT.Rows[0]["FrontConfigID"]), 1);

                SummarypFoxOrders(pFoxOrdersDT, pFoxSimpleDT, pFoxVitrinaDT, pFoxGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                   CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (pFlorencOrdersDT.Rows.Count > 0)
            {
                if (pFlorencOrdersDT.Rows.Count > 0)
                    FrontName = ProfileName(Convert.ToInt32(pFlorencOrdersDT.Rows[0]["FrontConfigID"]), 1);

                SummarypFlorenceOrders(pFlorencOrdersDT, pFlorencSimpleDT, pFlorencVitrinaDT, pFlorencGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                    CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            //if (Techno4MegaOrdersDT.Rows.Count > 0)
            //{
            //    FrontName = ProfileName(Convert.ToInt32(Techno4MegaOrdersDT.Rows[0]["FrontConfigID"]));
            //    SummaryTechno4MegaOrders(Techno4MegaOrdersDT, Techno4MegaDT, FrontName);
            //    OrdersToExcelSingly(ref hssfworkbook,
            //       CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR3);
            //    RowIndex++;
            //    RowIndex++;
            //}
            if (Techno5OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(Techno5OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryTechno5Orders(Techno5OrdersDT, Techno5SimpleDT, Techno5VitrinaDT, Techno5GridsDT, Techno5LuxDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                      CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR2, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            if (PR1OrdersDT.Rows.Count > 0)
            {
                DataTable DT = PR1OrdersDT.Clone();
                DataTable DT1 = PR1SimpleDT.Clone();
                DataTable DT2 = PR1VitrinaDT.Clone();
                DataTable DT3 = PR1GridsDT.Clone();
                DataRow[] rows = PR1OrdersDT.Select("FrontID=3631");
                foreach (DataRow item in rows)
                    DT.Rows.Add(item.ItemArray);
                rows = PR1SimpleDT.Select("FrontID=3631");
                foreach (DataRow item in rows)
                    DT1.Rows.Add(item.ItemArray);
                rows = PR1VitrinaDT.Select("FrontID=3631");
                foreach (DataRow item in rows)
                    DT2.Rows.Add(item.ItemArray);
                rows = PR1GridsDT.Select("FrontID=3631");
                foreach (DataRow item in rows)
                    DT3.Rows.Add(item.ItemArray);
                if (DT.Rows.Count > 0)
                {
                    FrontName = ProfileName(Convert.ToInt32(DT.Rows[0]["FrontConfigID"]), 1);
                    SummaryPR1Orders(DT, DT1, DT2, DT3, FrontName, ref AllSquare);
                    OrdersToExcelSingly(ref hssfworkbook,
                         CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, true, false, IsPR3, IsPRU8);
                    RowIndex++;
                    RowIndex++;
                }
                DT.Clear();
                DT1.Clear();
                DT2.Clear();
                DT3.Clear();
                rows = PR1OrdersDT.Select("FrontID=3632");
                foreach (DataRow item in rows)
                    DT.Rows.Add(item.ItemArray);
                rows = PR1SimpleDT.Select("FrontID=3632");
                foreach (DataRow item in rows)
                    DT1.Rows.Add(item.ItemArray);
                rows = PR1VitrinaDT.Select("FrontID=3632");
                foreach (DataRow item in rows)
                    DT2.Rows.Add(item.ItemArray);
                rows = PR1GridsDT.Select("FrontID=3632");
                foreach (DataRow item in rows)
                    DT3.Rows.Add(item.ItemArray);
                if (DT.Rows.Count > 0)
                {
                    FrontName = ProfileName(Convert.ToInt32(DT.Rows[0]["FrontConfigID"]), 1);
                    SummaryPR1Orders(DT, DT1, DT2, DT3, FrontName, ref AllSquare);
                    OrdersToExcelSingly(ref hssfworkbook,
                         CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, false, true, IsPR3, IsPRU8);
                    RowIndex++;
                    RowIndex++;
                }
            }
            if (PR3OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(PR3OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryTechno5Orders(PR3OrdersDT, PR3SimpleDT, PR3VitrinaDT, PR3GridsDT, PR3LuxDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                      CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, false, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            if (PRU8OrdersDT.Rows.Count > 0)
            {
                FrontName = ProfileName(Convert.ToInt32(PRU8OrdersDT.Rows[0]["FrontConfigID"]), 1);
                SummaryMarsel1Orders(PRU8OrdersDT, PRU8SimpleDT, PRU8VitrinaDT, PRU8GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook,
                     CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, false, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }
            AllSquare = decimal.Round(AllSquare, 3, MidpointRounding.AwayFromZero);
            OrdersToExcelSingly(ref hssfworkbook, CalibriBold11CS, CalibriBold11CS, ref sheet1, AllSquare, ref RowIndex);
        }

        private void OrdersToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, ref int RowIndex, bool IsPR1, bool IsPR2, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR2)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }
            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "Заказы");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 2), 1, "УТВЕРЖДАЮ_____________");
            //cell.CellStyle = Calibri11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            //cell.CellStyle = Calibri11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 4), 0, "Клиент:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 4), 1, ClientName);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 5), 0, "Партия:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 5), 1, BatchName);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 3), 0, "Задание №" + WorkAssignmentID.ToString());
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 3), 1, "Заказы");
            //cell.CellStyle = CalibriBold11CS;
            //RowIndex += 6;

            int ColumnIndex = -1;
            string ColumnName = string.Empty;

            for (int x = 0; x < SummOrdersDT.Columns.Count; x++)
            {
                if (SummOrdersDT.Columns[x].ColumnName == "Height" || SummOrdersDT.Columns[x].ColumnName == "Width"
                     || SummOrdersDT.Columns[x].ColumnName == "PR1Count" || SummOrdersDT.Columns[x].ColumnName == "PR2Count" || SummOrdersDT.Columns[x].ColumnName == "VitrinaCount")
                    continue;
                ColumnIndex++;
                ColumnName = SummOrdersDT.Columns[x].ColumnName;
                if (Contains(ColumnName, "_", StringComparison.OrdinalIgnoreCase))
                {
                    ColumnName = ColumnName.Substring(0, ColumnName.Length - 2);
                }
                if (ColumnName == "Sizes")
                {
                    ColumnName = "Размер";
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                    cell.CellStyle = TableHeaderCS;
                    //sheet1.SetColumnWidth(ColumnIndex, 12 * 256);
                    continue;
                }
                if (ColumnName == "TotalAmount")
                {
                    ColumnName = "Итого";
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                    cell.CellStyle = TableHeaderCS;
                    //sheet1.SetColumnWidth(ColumnIndex, 8 * 256);
                    continue;
                }
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                cell.CellStyle = TableHeaderCS;
                sheet1.SetColumnWidth(ColumnIndex, 19 * 256);
            }
            RowIndex++;
            TableHeaderCS.Alignment = HSSFCellStyle.ALIGN_LEFT;
            TableHeaderDecCS.Alignment = HSSFCellStyle.ALIGN_LEFT;

            HSSFFont FirstColF = hssfworkbook.CreateFont();
            FirstColF.FontHeightInPoints = 12;
            FirstColF.FontName = "MS Sans Serif";

            HSSFCellStyle FirstColCS = hssfworkbook.CreateCellStyle();
            FirstColCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            FirstColCS.LeftBorderColor = HSSFColor.BLACK.index;
            FirstColCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            FirstColCS.RightBorderColor = HSSFColor.BLACK.index;
            FirstColCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            FirstColCS.TopBorderColor = HSSFColor.BLACK.index;
            FirstColCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            FirstColCS.BottomBorderColor = HSSFColor.BLACK.index;
            FirstColCS.SetFont(FirstColF);
            for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
            {
                ColumnIndex = -1;
                for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
                {
                    if (SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width"
                     || SummOrdersDT.Columns[y].ColumnName == "PR1Count" || SummOrdersDT.Columns[y].ColumnName == "PR2Count" || SummOrdersDT.Columns[y].ColumnName == "VitrinaCount")
                        continue;
                    Type t = SummOrdersDT.Rows[x][y].GetType();

                    ColumnIndex++;

                    if (x == SummOrdersDT.Rows.Count - 1 && int.TryParse(SummOrdersDT.Rows[x][y].ToString(), out int IntValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(IntValue);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (x == SummOrdersDT.Rows.Count - 2 && double.TryParse(SummOrdersDT.Rows[x][y].ToString(), out double DecValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(DecValue);
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }

                    if (int.TryParse(SummOrdersDT.Rows[x][y].ToString(), out IntValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(IntValue);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(SummOrdersDT.Rows[x][y].ToString());
                        if (ColumnIndex == 0)
                            cell.CellStyle = FirstColCS;
                        else
                            cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
        }

        private void OrdersToExcelSingly(ref HSSFWorkbook hssfworkbook, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, decimal AllSquare, ref int RowIndex)
        {
            TableHeaderCS.Alignment = HSSFCellStyle.ALIGN_LEFT;
            TableHeaderDecCS.Alignment = HSSFCellStyle.ALIGN_LEFT;

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "ИТОГО:");
            cell.CellStyle = TableHeaderCS;

            cell = sheet1.CreateRow(RowIndex).CreateCell(1);
            cell.SetCellValue(Convert.ToDouble(AllSquare));
            cell.CellStyle = TableHeaderDecCS;
        }

        private void PR3HandsToExcel(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1, ref int RowIndex,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            if (DT.Rows.Count == 0)
                return;

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
            cell.CellStyle = CalibriBold15CS;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "MARTIN");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "терм");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "vitap");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "шкантование");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int AllTotalAmount = 0;
            int Height = 0;

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType"
                        || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType"
                            || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void PressToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal TotalSquare = 0;
            decimal AllTotalSquare = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 2, MidpointRounding.AwayFromZero);

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private string ProfileName(int ID)
        {
            string name = string.Empty;
            DataRow[] rows = ProfileNamesDT.Select("FrontConfigID=" + ID);
            if (rows.Count() > 0)
                name = rows[0]["TechStoreName"].ToString();
            return name;
        }

        private string ProfileName(int ID, int ProfileType)
        {
            //ProfileType
            string name = string.Empty;
            DataRow[] rows = ProfileNamesDT.Select("FrontConfigID=" + ID + " AND ProfileType=" + ProfileType);
            if (rows.Count() > 0)
                name = rows[0]["TechStoreName"].ToString();
            return name;
        }

        private void PRU8HandsToExcel(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1, ref int RowIndex,
                            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            if (DT.Rows.Count == 0)
                return;

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
            cell.CellStyle = CalibriBold15CS;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "MARTIN");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int AllTotalAmount = 0;
            int Height = 0;

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                        || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount" || RapidDT.Columns[y].ColumnName == "Notes")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (RapidDT.Columns[y].ColumnName == "ProfileType" || RapidDT.Columns[y].ColumnName == "ColorType" || RapidDT.Columns[y].ColumnName == "iCount"
                            || RapidDT.Columns[y].ColumnName == "PR1Count" || RapidDT.Columns[y].ColumnName == "PR2Count" || RapidDT.Columns[y].ColumnName == "VitrinaCount" || RapidDT.Columns[y].ColumnName == "Notes")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void RapidToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("MARTIN");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 6 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 16 * 256);
            sheet1.SetColumnWidth(6, 9 * 256);

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            RapidDT.Clear();
            Martin1Fronts(ref RapidDT);

            if (RapidDT.Rows.Count > 0)
            {
                MartinToExcel1(ref hssfworkbook, ref sheet1, ref RowIndex,
                    CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, RapidDT, WorkAssignmentID, DispatchDate, BatchName, ClientName);
            }

            RapidDT.Clear();
            Martin2Fronts(ref RapidDT);

            if (RapidDT.Rows.Count > 0)
            {
                RowIndex++;
                RowIndex++;
                MartinToExcel1(ref hssfworkbook, ref sheet1, ref RowIndex,
                    CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, RapidDT, WorkAssignmentID, DispatchDate, BatchName, ClientName);
            }

            RapidDT.Clear();
            MartinImpostFronts(ref RapidDT);

            if (RapidDT.Rows.Count > 0)
            {
                RowIndex++;
                RowIndex++;
                MartinToExcel1(ref hssfworkbook, ref sheet1, ref RowIndex,
                    CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, RapidDT, WorkAssignmentID, DispatchDate, BatchName, ClientName);
            }

            //RapidDT.Clear();
            //Martin2Fronts(ref RapidDT);

            //if (RapidDT.Rows.Count > 0)
            //{
            //    RowIndex++;
            //    RowIndex++;
            //    MartinToExcel2(ref hssfworkbook, ref sheet1, ref RowIndex,
            //        CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, RapidDT, WorkAssignmentID, DispatchDate, BatchName, ClientName);
            //}

            if (PR3OrdersDT.Rows.Count > 0)
            {
                RapidDT.Clear();
                MartinPR3Hands(PR3OrdersDT, ref RapidDT,
                    Convert.ToInt32(FrontMargins.Techno2Width), Convert.ToInt32(FrontMinSizes.PR3MinWidth), Convert.ToInt32(FrontMargins.Techno2Height), false);

                for (int i = 0; i < RapidDT.Rows.Count; i++)
                {
                    if (i == 0)
                        continue;
                    if (RapidDT.Rows[i]["Color"].ToString() == RapidDT.Rows[i - 1]["Color"].ToString())
                    {
                        RapidDT.Rows[i]["Color"] = string.Empty;
                    }
                }

                PR3HandsToExcel(ref hssfworkbook, ref sheet1, ref RowIndex,
                    CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, RapidDT, WorkAssignmentID, DispatchDate, BatchName, ClientName);
            }

            RowIndex++;
            if (PRU8OrdersDT.Rows.Count > 0)
            {
                RapidDT.Clear();
                MartinPRU8Hands(PRU8OrdersDT, ref RapidDT,
                    Convert.ToInt32(FrontMargins.Techno1Width), Convert.ToInt32(FrontMinSizes.Techno1MinWidth), Convert.ToInt32(FrontMargins.Techno1Height), false);
                PRU8HandsToExcel(ref hssfworkbook, ref sheet1, ref RowIndex,
                   CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, RapidDT, WorkAssignmentID, DispatchDate, BatchName, ClientName);
            }
        }

        private Tuple<bool, int, int, decimal, int> StandardImpostCount(int ID, int Height, int Width)
        {
            bool b = false;
            int Profile18 = 0;
            int HeightProfile18 = 0;
            int Profile16 = 0;
            decimal HeightProfile16 = 0;

            DataRow[] rows = StandardImpostDataTable.Select("FrontID=" + ID + " AND Height=" + Height + " AND Width=" + Width);
            if (rows.Count() > 0)
            {
                Profile18 = Convert.ToInt32(rows[0]["Profile18"]);
                HeightProfile18 = Convert.ToInt32(rows[0]["HeightProfile18"]);
                Profile16 = Convert.ToInt32(rows[0]["Profile16"]);
                HeightProfile16 = Convert.ToDecimal(rows[0]["HeightProfile16"]);
            }
            Tuple<bool, int, int, decimal, int> tuple = new Tuple<bool, int, int, decimal, int>(b, HeightProfile18, Profile18, HeightProfile16, Profile16);
            return tuple;
        }

        private void Stemas1ToExcelSingly(ref HSSFWorkbook hssfworkbook,
                    HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, ref int RowIndex, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (IsPR1)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPR3)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                cell.CellStyle = CalibriBold15CS;
            }
            if (IsPRU8)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                cell.CellStyle = CalibriBold15CS;
            }

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            int DisplayIndex = 0;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "18");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            if (bImpostMargin)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Смещ. имп.");
                cell.CellStyle = TableHeaderCS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "окл");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "фрез");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "сверл");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int CType = 0;
            int PType = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
                PType = Convert.ToInt32(DT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "IsBox")
                        continue;
                    if (!bImpostMargin && DT.Columns[y].ColumnName == "ImpostMargin")
                        continue;

                    if (bImpostMargin && DT.Columns[y].ColumnName == "ImpostMargin" && DT.Rows[x][y] != DBNull.Value)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = CalibriBold11CS;
                        continue;
                    }

                    if (Convert.ToBoolean(DT.Rows[x]["IsBox"]))
                    {
                        //HSSFCellStyle GreyCellStyle = hssfworkbook.CreateCellStyle();
                        //GreyCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
                        //GreyCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
                        //GreyCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
                        //GreyCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
                        //GreyCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
                        //GreyCellStyle.RightBorderColor = HSSFColor.BLACK.index;
                        //GreyCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
                        //GreyCellStyle.TopBorderColor = HSSFColor.BLACK.index;
                        //GreyCellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.BLACK.index;
                        //GreyCellStyle.FillPattern = HSSFCellStyle.THIN_FORWARD_DIAG;
                        //GreyCellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;

                        //cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        //cell.CellStyle = GreyCellStyle;
                        //cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        //cell.CellStyle = GreyCellStyle;
                    }

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1 && (PType != Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]) || CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "IsBox")
                            continue;
                        if (!bImpostMargin && DT.Columns[y].ColumnName == "ImpostMargin")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                    PType = Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]);
                    TotalAmount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "IsBox")
                            continue;
                        if (!bImpostMargin && DT.Columns[y].ColumnName == "ImpostMargin")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "IsBox")
                            continue;
                        if (!bImpostMargin && DT.Columns[y].ColumnName == "ImpostMargin")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void Stemas2ToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, ref int RowIndex, bool NeedHeader, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            HSSFCell cell = null;
            if (NeedHeader)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
                cell.CellStyle = Calibri11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
                cell.CellStyle = Calibri11CS;
                if (IsPR1)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-1 и ПР-2");
                    cell.CellStyle = CalibriBold15CS;
                }
                if (IsPR3)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПР-3");
                    cell.CellStyle = CalibriBold15CS;
                }
                if (IsPRU8)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "ПРУ-8");
                    cell.CellStyle = CalibriBold15CS;
                }

                if (DispatchDate.Length > 0)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                    cell.CellStyle = CalibriBold11CS;
                }

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
                cell.CellStyle = Calibri11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "16");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "шкантование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "терм");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int CType = 0;
            int PType = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
                PType = Convert.ToInt32(DT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ImpostMargin" || DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "IsBox")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1 && (PType != Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]) || CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ImpostMargin" || DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "IsBox")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                    PType = Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]);
                    TotalAmount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ImpostMargin" || DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "IsBox")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ImpostMargin" || DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "IsBox")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void StemasFlorenceProfil16(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int WidthMargin, int WidthMin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistWidthTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID = -1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID = -1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["IsBox"] = false;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }

            DT1.Clear();
            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID <> -1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["IsBox"] = false;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }
        
        private void StemasAllFrontsProfil16(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int WidthMargin, int WidthMin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistWidthTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                //DV.RowFilter = "TechnoColorID = -1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["IsBox"] = false;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasFlorenceProfil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) + " AND ImpostMargin=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    int ImpostMargin = Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= HeightMargin)
                    {
                        Height = HeightMargin;
                        IsBox = true;
                    }

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType + " AND ImpostMargin=" + ImpostMargin);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }

                        if (bImpostMargin && ImpostMargin != 0)
                            NewRow["ImpostMargin"] = ImpostMargin;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
            DT1.Clear();
            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID<>-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID<>-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) + " AND ImpostMargin=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    int ImpostMargin = Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= HeightMargin)
                    {
                        Height = HeightMargin;
                        IsBox = true;
                    }

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType + " AND ImpostMargin=" + ImpostMargin);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }

                        if (bImpostMargin && ImpostMargin != 0)
                            NewRow["ImpostMargin"] = ImpostMargin;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }
        
        private void StemasAllFrontsProfil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                //DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) + " AND ImpostMargin=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    int ImpostMargin = Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= HeightMargin)
                    {
                        Height = HeightMargin;
                        IsBox = true;
                    }

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType + " AND ImpostMargin=" + ImpostMargin);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }

                        if (bImpostMargin && ImpostMargin != 0)
                            NewRow["ImpostMargin"] = ImpostMargin;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasFrontsByHeight(ref DataTable DestinationDT)
        {
            int ProfileType = 1;
            if (Marsel1OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(Marsel1OrdersDT, ref DestinationDT, "Марсель П-018", Convert.ToInt32(FrontMargins.Marsel1Height), ProfileType++, true);
            if (Marsel5OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(Marsel5OrdersDT, ref DestinationDT, "Марсель-5 П-071", Convert.ToInt32(FrontMargins.Marsel5Height), ProfileType++, true);
            if (PortoOrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(PortoOrdersDT, ref DestinationDT, "Порто П-0111", Convert.ToInt32(FrontMargins.PortoHeight), ProfileType++, true);
            if (MonteOrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(MonteOrdersDT, ref DestinationDT, "Монте П-0112", Convert.ToInt32(FrontMargins.MonteHeight), ProfileType++, true);
            if (Marsel3OrdersDT.Rows.Count > 0)
                StemasMarselProfil18(Marsel3OrdersDT, ref DestinationDT, "Марсель-3 П-041", Convert.ToInt32(FrontMargins.Marsel3Height), ProfileType++, true);
            if (Marsel4OrdersDT.Rows.Count > 0)
                StemasMarsel4Profil18(Marsel4OrdersDT, ref DestinationDT, "Марсель-4 П-066", Convert.ToInt32(FrontMargins.Marsel4Height), Convert.ToInt32(FrontMargins.Marsel4Height1), ProfileType++, true);
            if (Jersy110OrdersDT.Rows.Count > 0)
                StemasMarsel4Profil18(Jersy110OrdersDT, ref DestinationDT, "Джерси П-0110", Convert.ToInt32(FrontMargins.Jersy110Height), Convert.ToInt32(FrontMargins.Jersy110Height1), ProfileType++, true);
            if (Techno1OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(Techno1OrdersDT, ref DestinationDT, "Техно П-106", Convert.ToInt32(FrontMargins.Techno1Height), ProfileType++, true);
            if (Techno2OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(Techno2OrdersDT, ref DestinationDT, "Техно-2 П-206", Convert.ToInt32(FrontMargins.Techno2Height), ProfileType++, true);

            if (Techno4OrdersDT.Rows.Count > 0)
                StemasTechno4Profil18(Techno4OrdersDT, ref DestinationDT, "Техно-4 П-406", Convert.ToInt32(FrontMargins.Techno4Height),
                    Convert.ToInt32(FrontMargins.Techno4NarrowWidth), ProfileType++, true);
            if (pFoxOrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(pFoxOrdersDT, ref DestinationDT, "Фокс П-042-4", Convert.ToInt32(FrontMargins.pFoxHeight),
                    ProfileType++, true);
            if (pFlorencOrdersDT.Rows.Count > 0)
                StemasFlorenceProfil18(pFlorencOrdersDT, ref DestinationDT, "П-01418", Convert.ToInt32(FrontMargins.pFlorencHeight),
                    ProfileType++, true);

            if (Techno5OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(Techno5OrdersDT, ref DestinationDT, "Техно-5 П-406", Convert.ToInt32(FrontMargins.Techno5Height), ProfileType++, true);

            if (PR1OrdersDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = PR1OrdersDT.Copy();
                for (int i = 0; i < PR1OrdersDT.Rows.Count; i++)
                {
                    object x1 = PR1OrdersDT.Rows[i]["Height"];
                    object x2 = PR1OrdersDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                StemasAllFrontsProfil18(TempPR1OrdersDT, ref DestinationDT, "Марсель-3 П-041", Convert.ToInt32(FrontMargins.Marsel3Height), ProfileType++, true);
                //StemasImpostProfil18(TempPR1OrdersDT, ref DestinationDT, "Марсель-3 П-041", Convert.ToInt32(FrontMargins.Marsel3Height), ProfileType++, true);
            }
            if (ShervudOrdersDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = ShervudOrdersDT.Copy();
                for (int i = 0; i < ShervudOrdersDT.Rows.Count; i++)
                {
                    object x1 = ShervudOrdersDT.Rows[i]["Height"];
                    object x2 = ShervudOrdersDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                StemasAllFrontsProfil18(TempPR1OrdersDT, ref DestinationDT, "Шервуд П-043-4", Convert.ToInt32(FrontMargins.ShervudHeight), ProfileType++, true);
            }
            if (PR3OrdersDT.Rows.Count > 0)
                StemasPR3Profil18(PR3OrdersDT, ref DestinationDT, "Техно-2 П-206", Convert.ToInt32(FrontMargins.Techno2Height), ProfileType++, true);
            if (PRU8OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil18(PRU8OrdersDT, ref DestinationDT, "Техно П-106", Convert.ToInt32(FrontMargins.Techno1Height), ProfileType++, true);
            //if (Marsel3OrdersDT.Rows.Count > 0)
            //    StemasImpostProfil18(Marsel3OrdersDT, ref DestinationDT, "Марсель-3 П-041", Convert.ToInt32(FrontMargins.Marsel3Height), ProfileType++, true);
            //if (Marsel4OrdersDT.Rows.Count > 0)
            //    StemasMarsel4ImpostProfil18(Marsel4OrdersDT, ref DestinationDT, "Марсель-4 П-066", Convert.ToInt32(FrontMargins.Marsel4Height), Convert.ToInt32(FrontMargins.Marsel4Height1), ProfileType++, true);
        }

        private void StemasFrontsByWidth(ref DataTable DestinationDT)
        {
            int ProfileType = 1;
            if (Marsel1OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(Marsel1OrdersDT, ref DestinationDT, "Марсель П-016",
                    Convert.ToInt32(FrontMargins.Marsel1Width), Convert.ToInt32(FrontMinSizes.Marsel1MinWidth), ProfileType++, true);
            if (Marsel5OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(Marsel5OrdersDT, ref DestinationDT, "Марсель-5 П-171",
                    Convert.ToInt32(FrontMargins.Marsel5Width), Convert.ToInt32(FrontMinSizes.Marsel5MinWidth), ProfileType++, true);
            if (PortoOrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(PortoOrdersDT, ref DestinationDT, "Порто П-111",
                    Convert.ToInt32(FrontMargins.PortoWidth), Convert.ToInt32(FrontMinSizes.PortoMinWidth), ProfileType++, true);
            if (MonteOrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(MonteOrdersDT, ref DestinationDT, "Монте П-112",
                    Convert.ToInt32(FrontMargins.MonteWidth), Convert.ToInt32(FrontMinSizes.MonteMinWidth), ProfileType++, true);
            if (Marsel3OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(Marsel3OrdersDT, ref DestinationDT, "Марсель-3 П-141",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), ProfileType++, true);
            if (Marsel4OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(Marsel4OrdersDT, ref DestinationDT, "Марсель-4 П-166",
                    Convert.ToInt32(FrontMargins.Marsel4Width), Convert.ToInt32(FrontMinSizes.Marsel4MinWidth), ProfileType++, true);
            if (Jersy110OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(Jersy110OrdersDT, ref DestinationDT, "Джерси П-110",
                    Convert.ToInt32(FrontMargins.Jersy110Width), Convert.ToInt32(FrontMinSizes.Jersy110MinWidth), ProfileType++, true);
            if (Techno1OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(Techno1OrdersDT, ref DestinationDT, "Техно П-116",
                    Convert.ToInt32(FrontMargins.Techno1Width), Convert.ToInt32(FrontMinSizes.Techno1MinWidth), ProfileType++, true);
            if (Techno2OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(Techno2OrdersDT, ref DestinationDT, "Техно-2 П-216",
                    Convert.ToInt32(FrontMargins.Techno2Width), Convert.ToInt32(FrontMinSizes.Techno2MinWidth), ProfileType++, true);

            //DataTable DT = Techno4MegaOrdersDT.Clone();
            //foreach (DataRow item in Techno4OrdersDT.Rows)
            //    DT.Rows.Add(item.ItemArray);
            //foreach (DataRow item in Techno4MegaOrdersDT.Rows)
            //    DT.Rows.Add(item.ItemArray);
            if (Techno4OrdersDT.Rows.Count > 0)
                StemasTechno4Profil16(Techno4OrdersDT, ref DestinationDT, "Техно-4 П-416",
                    Convert.ToInt32(FrontMargins.Techno4Width), Convert.ToInt32(FrontMinSizes.Techno4MinWidth),
                    Convert.ToInt32(FrontMargins.Techno4NarrowHeight), ProfileType++, true);
            if (pFoxOrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(pFoxOrdersDT, ref DestinationDT, "Фокс П-042-4",
                    Convert.ToInt32(FrontMargins.pFoxWidth), Convert.ToInt32(FrontMinSizes.pFoxMinWidth), ProfileType++, true);
            if (pFlorencOrdersDT.Rows.Count > 0)
            {
                StemasFlorenceProfil16(pFlorencOrdersDT, ref DestinationDT, "П-1418-0",
                    Convert.ToInt32(FrontMargins.pFlorencWidth), Convert.ToInt32(FrontMinSizes.pFlorencMinWidth), ProfileType++, true);
                StemasImpost(pFlorencOrdersDT, ref DestinationDT, ProfileType++, true);
            }

            if (Techno5OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(Techno5OrdersDT, ref DestinationDT, "Техно-5 П-516",
                    Convert.ToInt32(FrontMargins.Techno5Width), Convert.ToInt32(FrontMinSizes.Techno5MinWidth), ProfileType++, true);

            if (PR1OrdersDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = PR1OrdersDT.Copy();
                for (int i = 0; i < PR1OrdersDT.Rows.Count; i++)
                {
                    object x1 = PR1OrdersDT.Rows[i]["Height"];
                    object x2 = PR1OrdersDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                StemasAllFrontsProfil16(TempPR1OrdersDT, ref DestinationDT, "Марсель-3 П-141",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), ProfileType++, true);
                StemasImpostProfil16(TempPR1OrdersDT, ref DestinationDT, "Марсель-3 П-141",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), ProfileType++, true);
            }
            if (ShervudOrdersDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = ShervudOrdersDT.Copy();
                for (int i = 0; i < ShervudOrdersDT.Rows.Count; i++)
                {
                    object x1 = ShervudOrdersDT.Rows[i]["Height"];
                    object x2 = ShervudOrdersDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                StemasAllFrontsProfil16(TempPR1OrdersDT, ref DestinationDT, "Шервуд П-042-4",
                    Convert.ToInt32(FrontMargins.ShervudWidth), Convert.ToInt32(FrontMinSizes.ShervudMinWidth), ProfileType++, true);
            }
            if (PR3OrdersDT.Rows.Count > 0)
                StemasPR3Profil16(PR3OrdersDT, ref DestinationDT, "Техно-2 П-216",
                    Convert.ToInt32(FrontMargins.Techno2Width), Convert.ToInt32(FrontMinSizes.PR3MinWidth), ProfileType++, true);
            if (PRU8OrdersDT.Rows.Count > 0)
                StemasAllFrontsProfil16(PRU8OrdersDT, ref DestinationDT, "Техно П-116",
                    Convert.ToInt32(FrontMargins.Techno1Width), Convert.ToInt32(FrontMinSizes.Techno1MinWidth), ProfileType++, true);
            if (Marsel3OrdersDT.Rows.Count > 0)
                StemasImpostProfil16(Marsel3OrdersDT, ref DestinationDT, "Марсель-3 П-141",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), ProfileType++, true);
            if (Marsel4OrdersDT.Rows.Count > 0)
                StemasImpostProfil16(Marsel4OrdersDT, ref DestinationDT, "Марсель-4 П-166",
                    Convert.ToInt32(FrontMargins.Marsel4Width), Convert.ToInt32(FrontMinSizes.Marsel4MinWidth), ProfileType++, true);
        }

        private void StemasImpost(DataTable SourceDT, ref DataTable DestinationDT, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistSizesTable(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }
            DT1 = OrderedTechnoFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) + " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        int Width = Convert.ToInt32(DT2.Rows[j]["Width"]);
                        Tuple<bool, int, int, decimal, int> tuple = StandardImpostCount(Convert.ToInt32(Srows[0]["FrontID"]), Height, Width);

                        bool b = tuple.Item1;
                        int HeightProfile18 = tuple.Item2;
                        int Profile18 = tuple.Item3;
                        decimal HeightProfile16 = tuple.Item4;
                        int Profile16 = tuple.Item5;

                        //Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontConfigID"]), Convert.ToInt32(SourceDT.Rows[0]["TechnoProfileID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            //if (Convert.ToInt32(item["InsetTypeID"]) != 1 && (Convert.ToInt32(item["Height"]) - 10 < WidthMargin || Convert.ToInt32(item["Width"]) - 10 < WidthMargin))
                            //    BoxCount += Convert.ToInt32(item["Count"]);
                            //if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                            //    VitrinaCount += Convert.ToInt32(item["Count"]);

                            Count += Convert.ToInt32(item["Count"]);
                        }

                        string ProfileName1 = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 2);
                        {
                            DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + Color + "' AND Height=" + HeightProfile18);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Front"] = ProfileName1;
                                NewRow["Color"] = Color;
                                NewRow["Height"] = HeightProfile18;
                                NewRow["Count"] = Count * Profile18;
                                NewRow["ProfileType"] = ProfileType;
                                NewRow["IsBox"] = false;
                                NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * Profile18;
                            }
                        }
                        {
                            var filteredRows = DestinationDT
                                .AsEnumerable()
                                .Where(row => row.Field<decimal>("Height") == HeightProfile16
                                && row.Field<string>("Front") == ProfileName1
                                && row.Field<string>("Color") == Color);

                            DataRow[] rows = filteredRows.ToArray();
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Front"] = ProfileName1;
                                NewRow["Color"] = Color;
                                NewRow["Height"] = HeightProfile16;
                                NewRow["Count"] = Count * Profile16;
                                NewRow["ProfileType"] = ProfileType;
                                NewRow["IsBox"] = false;
                                NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * Profile16;
                            }
                        }
                    }
                }
            }
        }

        private void StemasImpostProfil16(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int WidthMargin, int WidthMin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistWidthTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 2) + " ИМПОСТ";
                            NewRow["Color"] = FrameColor;
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["IsBox"] = false;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                    }
                }
            }
        }

        private void StemasImpostProfil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) + " AND ImpostMargin=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    int ImpostMargin = Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= HeightMargin)
                    {
                        Height = HeightMargin;
                        IsBox = true;
                    }

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType + " AND ImpostMargin=" + ImpostMargin);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1) + " ИМПОСТ";
                            NewRow["Color"] = FrameColor;
                        }
                        if (bImpostMargin && ImpostMargin != 0)
                            NewRow["ImpostMargin"] = ImpostMargin;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasMarsel4ImpostProfil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int HeightMargin1, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) + " AND ImpostMargin=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    int ImpostMargin = Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height > HeightMargin1)
                    {
                    }
                    else
                    {
                        if (Height <= HeightMargin + 1)
                            Height = HeightMargin;
                        if (Height > HeightMargin + 1 && Height <= HeightMargin1)
                            Height = HeightMargin1;
                        IsBox = true;
                    }
                    //if (Height <= HeightMargin + 1)
                    //{
                    //    Height = HeightMargin;
                    //    IsBox = true;
                    //}
                    //if (Height < HeightMargin + 1 && Height >= HeightMargin1)
                    //{
                    //    Height = HeightMargin1;
                    //    IsBox = true;
                    //}

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType + " AND ImpostMargin=" + ImpostMargin);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1) + " ИМПОСТ";
                            NewRow["Color"] = FrameColor;
                        }
                        if (bImpostMargin && ImpostMargin != 0)
                            NewRow["ImpostMargin"] = ImpostMargin;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasMarsel4Profil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int HeightMargin1, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) + " AND ImpostMargin=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    int ImpostMargin = Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height > HeightMargin1)
                    {
                    }
                    else
                    {
                        if (Height <= HeightMargin + 1)
                            Height = HeightMargin;
                        if (Height > HeightMargin + 1 && Height <= HeightMargin1)
                            Height = HeightMargin1;
                        IsBox = true;
                    }
                    //if (Height <= HeightMargin + 1)
                    //{
                    //    Height = HeightMargin;
                    //    IsBox = true;
                    //}
                    //if (Height < HeightMargin + 1 && Height >= HeightMargin1)
                    //{
                    //    Height = HeightMargin1;
                    //    IsBox = true;
                    //}

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType + " AND ImpostMargin=" + ImpostMargin);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        if (bImpostMargin && ImpostMargin != 0)
                            NewRow["ImpostMargin"] = ImpostMargin;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasMarselProfil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=-1 AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) + " AND ImpostMargin=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    int ImpostMargin = Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    if (Height <= HeightMargin)
                    {
                        Height = HeightMargin;
                        IsBox = true;
                    }

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType + " AND ImpostMargin=" + ImpostMargin);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        if (bImpostMargin && ImpostMargin != 0)
                            NewRow["ImpostMargin"] = ImpostMargin;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasPR3Profil16(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int WidthMargin, int WidthMin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistWidthTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 3;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 3;
                    }
                }
            }
        }

        private void StemasPR3Profil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= HeightMargin)
                    {
                        Height = HeightMargin;
                        IsBox = true;
                    }

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) * 2;
                    }
                }
            }
        }

        private void StemasShervudProfil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                //DV.RowFilter = "TechnoColorID=-1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) + " AND ImpostMargin=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    int ImpostMargin = Convert.ToInt32(DistinctSizesDT.Rows[j]["ImpostMargin"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= HeightMargin)
                    {
                        Height = HeightMargin;
                        IsBox = true;
                    }

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType + " AND ImpostMargin=" + ImpostMargin);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        if (bImpostMargin && ImpostMargin != 0)
                            NewRow["ImpostMargin"] = ImpostMargin;
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasTechno4Profil16(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int WidthMargin, int WidthMin, int HeightNarrowMargin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistWidthTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height < WidthMargin)
                        Height = Height - HeightNarrowMargin;
                    else
                        Height = Height - WidthMargin;
                    if (Height <= WidthMin)
                        Height = WidthMin;

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height + " AND ProfileType=" + ProfileType);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasTechno4Profil18(DataTable SourceDT, ref DataTable DestinationDT, string Front,
            int HeightMargin, int WidthNarrowMargin, int ProfileType, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DistinctSizesDT = DistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                bool AlreadyExist = false;
                for (int j = 0; j < DistinctSizesDT.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    bool IsBox = false;
                    int Count = 0;
                    int Height = Convert.ToInt32(DistinctSizesDT.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    //if (Height < HeightMargin)
                    //    Height = Height - WidthNarrowMargin;
                    if (Height <= HeightMargin)
                    {
                        //Height = HeightMargin;
                        IsBox = true;
                    }

                    DataRow[] rows = DestinationDT.Select("Color='" + FrameColor + "' AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        if (!AlreadyExist)
                        {
                            AlreadyExist = true;
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]), 1);
                            NewRow["Color"] = FrameColor;
                        }
                        NewRow["Height"] = Height;
                        NewRow["Count"] = Count * 2;
                        NewRow["ProfileType"] = ProfileType;
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["IsBox"] = IsBox;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                    }
                }
            }
        }

        private void StemasToExcel(ref HSSFWorkbook hssfworkbook,
                                                                                                                                    HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, bool IsPR1, bool IsPR3, bool IsPRU8)
        {
            StemasDT.Clear();
            StemasFrontsByHeight(ref StemasDT);
            if (StemasDT.Rows.Count == 0)
                return;

            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Stemas");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 6 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 11 * 256);
            sheet1.SetColumnWidth(5, 11 * 256);
            sheet1.SetColumnWidth(6, 11 * 256);
            sheet1.SetColumnWidth(7, 11 * 256);

            DataTable DT = StemasDT.Copy();
            DataColumn Col1 = new DataColumn();
            DataColumn Col2 = new DataColumn();
            DataColumn Col3 = new DataColumn();

            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
            Col3 = DT.Columns.Add("Col3", System.Type.GetType("System.String"));

            if (bImpostMargin)
            {
                Col1.SetOrdinal(5);
                Col2.SetOrdinal(6);
                Col3.SetOrdinal(7);
            }
            else
            {
                Col1.SetOrdinal(4);
                Col2.SetOrdinal(5);
                Col3.SetOrdinal(6);
            }
            //DT.Columns["IsBox"].SetOrdinal(8);

            if (DT.Rows.Count > 0)
            {
                Stemas1ToExcelSingly(ref hssfworkbook,
                       CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            StemasDT.Clear();
            StemasFrontsByWidth(ref StemasDT);

            DT.Dispose();
            Col1.Dispose();
            Col2.Dispose();
            Col3.Dispose();
            DT = StemasDT.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);
            Col2.SetOrdinal(5);
            DT.Columns["IsBox"].SetOrdinal(6);

            if (DT.Rows.Count > 0)
            {
                Stemas2ToExcelSingly(ref hssfworkbook,
                      CalibriBold15CS, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex, false, IsPR1, IsPR3, IsPRU8);
                RowIndex++;
                RowIndex++;
            }

            ComecDT.Clear();
            CombineComecSimple(ref ComecDT);

            DT = new DataTable();
            DT = ComecDT.Copy();

            if (DT.Rows.Count > 0)
            {
                ComecToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, "Comec", DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            RapidTechnoInsetDT.Clear();
            RapidTechnoInsetSimple(ref RapidTechnoInsetDT);

            if (RapidTechnoInsetDT.Rows.Count > 0)
            {
                RapidTechnoInsetToExcel(ref hssfworkbook, ref sheet1,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, RapidTechnoInsetDT, WorkAssignmentID, "Вставка-2", DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
        }

        private void SummaryMarsel1Orders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryMarsel3Orders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, true);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, true);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, true);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryMarsel5Orders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, DataTable LayoutDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            CollectOrders(DistinctSizesDT, LayoutDT, ref SummOrdersDT, 4, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummarypFlorenceOrders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }
        
        private void SummarypFoxOrders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryPR1Orders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryShervudOrders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryTechno1Orders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, DataTable LuxDT, DataTable MegaDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            CollectOrders(DistinctSizesDT, LuxDT, ref SummOrdersDT, 4, FrontName, false);
            CollectOrders(DistinctSizesDT, MegaDT, ref SummOrdersDT, 5, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryTechno2Orders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, DataTable LuxDT, DataTable MegaDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            CollectOrders(DistinctSizesDT, LuxDT, ref SummOrdersDT, 4, FrontName, false);
            CollectOrders(DistinctSizesDT, MegaDT, ref SummOrdersDT, 5, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryTechno4Orders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, DataTable LuxDT, DataTable MegaDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            CollectOrders(DistinctSizesDT, LuxDT, ref SummOrdersDT, 4, FrontName, false);
            CollectOrders(DistinctSizesDT, MegaDT, ref SummOrdersDT, 5, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryTechno5Orders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, DataTable LuxDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName, false);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName, false);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName, false);
            CollectOrders(DistinctSizesDT, LuxDT, ref SummOrdersDT, 5, FrontName, false);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void Techno4AllInsets(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightProfilMargin, int WidthProfilMargin, int HeightNarrowMargin, int WidthNarrowMargin, int HeightMargin, int WidthMargin,
            int HeightMin, int WidthMin, bool NeedSwap, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]);
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]);
                            int TempHeightMargin = HeightMargin;
                            int TempWidthMargin = WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            if (Height < HeightProfilMargin)
                                TempHeightMargin = HeightNarrowMargin;
                            if (Width < WidthProfilMargin)
                                TempWidthMargin = WidthNarrowMargin;

                            Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - TempHeightMargin;
                            Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - TempWidthMargin;

                            string InsetColor = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            int GroupID = Convert.ToInt32(InsetTypesDataTable.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]))[0]["GroupID"]);
                            if (GroupID == 7 || GroupID == 8)
                            {
                                InsetColor = InsetColor.Insert(0, "фил ");
                                if ((Height > HeightMin && Width > WidthMin))
                                    continue;
                            }

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            //if (TechnoInsetColorID == 128)
                            //{
                            //    InsetColor = "мега " + InsetColor + " вит";
                            //}

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                if (NeedSwap)
                                {
                                    NewRow["Width"] = Height;
                                    NewRow["Height"] = Width;
                                }
                                else
                                {
                                    NewRow["Height"] = Height;
                                    NewRow["Width"] = Width;
                                }
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void Techno4InsetsGridsOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightProfilMargin, int WidthProfilMargin, int HeightNarrowMargin, int WidthNarrowMargin,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]);
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]);
                            int TempHeightMargin = HeightMargin;
                            int TempWidthMargin = WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            if (Height < HeightProfilMargin)
                                TempHeightMargin = HeightNarrowMargin;
                            if (Width < WidthProfilMargin)
                                TempWidthMargin = WidthNarrowMargin;

                            Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - TempHeightMargin;
                            Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - TempWidthMargin;

                            if (Height < 10 || Width < 10)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            string Name = string.Empty;
                            if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 685 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 688 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 29470)
                                Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " 45 " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 686 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 687 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 29471)
                                Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " 90 " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = Name;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void Techno4LuxOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightProfilMargin, int WidthProfilMargin, int HeightNarrowMargin, int WidthNarrowMargin,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]);
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]);
                            int TempHeightMargin = HeightMargin;
                            int TempWidthMargin = WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            if (Height < HeightProfilMargin)
                                TempHeightMargin = HeightNarrowMargin;
                            if (Width < WidthProfilMargin)
                                TempWidthMargin = WidthNarrowMargin;

                            Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - TempHeightMargin;
                            Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - TempWidthMargin;

                            int LuxCount = 0;
                            string InsetColor = "люкс " + GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            LuxCount = Convert.ToInt32(Math.Truncate(Height / 65m));
                            if (LuxCount == 0)
                                LuxCount = 1;
                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["GlassCount"] = LuxCount * Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                                rows[0]["GlassCount"] = Convert.ToInt32(rows[0]["GlassCount"]) + LuxCount * Count;
                            }
                        }
                    }
                }
            }
        }

        private void Techno4MegaOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int HeightProfilMargin, int WidthProfilMargin, int HeightNarrowMargin, int WidthNarrowMargin,
            int HeightMargin, int WidthMargin, int HeightMin, int WidthMin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetTypeID, TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetTypeID", "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]);
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]);
                            int TempHeightMargin = HeightMargin;
                            int TempWidthMargin = WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            if (Height < HeightProfilMargin)
                                TempHeightMargin = HeightNarrowMargin;
                            if (Width < WidthProfilMargin)
                                TempWidthMargin = WidthNarrowMargin;

                            Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - TempHeightMargin;
                            Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - TempWidthMargin;

                            int GlassCount = 0;
                            int MegaCount = 0;
                            string InsetColor = "мега " + GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            foreach (DataRow item in Srows)
                            {
                                Count += Convert.ToInt32(item["Count"]);
                            }

                            if (Height <= HeightMin)
                                Height = HeightMin;
                            if (Width <= WidthMin)
                                Width = WidthMin;

                            GetMegaInsetStickCount(Height, ref GlassCount, ref MegaCount);

                            if (TechnoInsetColorID == 3943)
                            {
                                GlassCount = 0;
                                InsetColor += " вит";
                            }
                            else
                                InsetColor += "/" + GetInsetColorName(TechnoInsetColorID);

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["GlassCount"] = GlassCount * Count;
                                NewRow["MegaCount"] = MegaCount * Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                                rows[0]["GlassCount"] = Convert.ToInt32(rows[0]["GlassCount"]) + GlassCount * Count;
                                rows[0]["MegaCount"] = Convert.ToInt32(rows[0]["MegaCount"]) + MegaCount * Count;
                            }
                        }
                    }
                }
            }
        }

        private void TotalInfoToExcel(ref HSSFWorkbook hssfworkbook,
                                                                                                                            HSSFCellStyle CalibriBold15CS, HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            TotalInfoDT.Clear();
            U(ref TotalInfoDT);

            if (TotalInfoDT.Rows.Count == 0)
                return;

            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Общая информация");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 6 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 16 * 256);

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "Сводка по заданию");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            for (int x = 0; x < TotalInfoDT.Rows.Count; x++)
            {
                for (int y = 0; y < TotalInfoDT.Columns.Count; y++)
                {
                    Type t = TotalInfoDT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(TotalInfoDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(TotalInfoDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(TotalInfoDT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                RowIndex++;
            }
        }

        private void TotalSum(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2, int WidthMargin, int WidthMin, int HeightMargin)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    decimal SticksCount = 0;
                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    SticksCount = (Height + 4) * Count * 2;
                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "'");
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["SticksCount"] = SticksCount;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                        rows[0]["SticksCount"] = Convert.ToDecimal(rows[0]["SticksCount"]) + SticksCount;
                }

                SizesASC = "Width ASC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    decimal SticksCount = 0;
                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    SticksCount = (Height + 4) * Count * 2;
                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "'");
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["SticksCount"] = SticksCount;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                        rows[0]["SticksCount"] = Convert.ToDecimal(rows[0]["SticksCount"]) + SticksCount;
                }
            }
        }

        private void TotalSumTechno4(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, int HeightNarrowMargin)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    decimal SticksCount = 0;
                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    SticksCount = (Height + 4) * Count * 2;
                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "'");
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["SticksCount"] = SticksCount;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                        rows[0]["SticksCount"] = Convert.ToDecimal(rows[0]["SticksCount"]) + SticksCount;
                }

                SizesASC = "Width ASC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    decimal SticksCount = 0;
                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height < HeightMargin)
                        Height = Height - HeightNarrowMargin;
                    else
                        Height = Height - WidthMargin;
                    if (Height <= WidthMin)
                        Height = WidthMin;

                    SticksCount = (Height + 4) * Count * 2;
                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "'");
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["SticksCount"] = SticksCount;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                        rows[0]["SticksCount"] = Convert.ToDecimal(rows[0]["SticksCount"]) + SticksCount;
                }
            }
        }

        private void U(ref DataTable DestinationDT)
        {
            if (Marsel1OrdersDT.Rows.Count > 0)
            {
                TotalSum(Marsel1OrdersDT, ref DestinationDT, "Марсель П-016", "Марсель П-018",
                    Convert.ToInt32(FrontMargins.Marsel1Width), Convert.ToInt32(FrontMinSizes.Marsel1MinWidth), Convert.ToInt32(FrontMargins.Marsel1Height));
            }
            if (Marsel5OrdersDT.Rows.Count > 0)
            {
                TotalSum(Marsel5OrdersDT, ref DestinationDT, "Марсель-5 П-171", "Марсель-5 П-071",
                    Convert.ToInt32(FrontMargins.Marsel5Width), Convert.ToInt32(FrontMinSizes.Marsel5MinWidth), Convert.ToInt32(FrontMargins.Marsel5Height));
            }
            if (PortoOrdersDT.Rows.Count > 0)
            {
                TotalSum(PortoOrdersDT, ref DestinationDT, "Порто П-111", "Порто П-0111",
                    Convert.ToInt32(FrontMargins.PortoWidth), Convert.ToInt32(FrontMinSizes.PortoMinWidth), Convert.ToInt32(FrontMargins.PortoHeight));
            }
            if (MonteOrdersDT.Rows.Count > 0)
            {
                TotalSum(MonteOrdersDT, ref DestinationDT, "Монте П-112", "Монте П-0112",
                    Convert.ToInt32(FrontMargins.MonteWidth), Convert.ToInt32(FrontMinSizes.MonteMinWidth), Convert.ToInt32(FrontMargins.MonteHeight));
            }
            if (Marsel3OrdersDT.Rows.Count > 0)
            {
                TotalSum(Marsel3OrdersDT, ref DestinationDT, "Марсель-3 П-141", "Марсель-3 П-041",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), Convert.ToInt32(FrontMargins.Marsel3Height));
            }
            if (Marsel4OrdersDT.Rows.Count > 0)
            {
                TotalSum(Marsel4OrdersDT, ref DestinationDT, "Марсель-4 П-166", "Марсель-4 П-066",
                    Convert.ToInt32(FrontMargins.Marsel4Width), Convert.ToInt32(FrontMinSizes.Marsel4MinWidth), Convert.ToInt32(FrontMargins.Marsel4Height));
            }
            if (Jersy110OrdersDT.Rows.Count > 0)
            {
                TotalSum(Jersy110OrdersDT, ref DestinationDT, "Джерси П-110", "Джерси П-0110",
                    Convert.ToInt32(FrontMargins.Jersy110Width), Convert.ToInt32(FrontMinSizes.Jersy110MinWidth), Convert.ToInt32(FrontMargins.Jersy110Height));
            }
            if (Techno1OrdersDT.Rows.Count > 0)
            {
                TotalSum(Techno1OrdersDT, ref DestinationDT, "Техно П-116", "Техно П-106",
                    Convert.ToInt32(FrontMargins.Techno1Width), Convert.ToInt32(FrontMinSizes.Techno1MinWidth), Convert.ToInt32(FrontMargins.Techno1Height));
            }
            if (Techno2OrdersDT.Rows.Count > 0)
            {
                TotalSum(Techno2OrdersDT, ref DestinationDT, "Техно-2 П-216", "Техно-2 П-206",
                    Convert.ToInt32(FrontMargins.Techno2Width), Convert.ToInt32(FrontMinSizes.Techno2MinWidth), Convert.ToInt32(FrontMargins.Techno2Height));
            }

            if (Techno4OrdersDT.Rows.Count > 0)
                TotalSumTechno4(Techno4OrdersDT, ref DestinationDT, "Техно-4 П-416", "Техно-4 П-406",
                    Convert.ToInt32(FrontMargins.Techno4Width), Convert.ToInt32(FrontMinSizes.Techno4MinWidth), Convert.ToInt32(FrontMargins.Techno4Height),
                    Convert.ToInt32(FrontMargins.Techno4NarrowHeight));
            if (pFoxOrdersDT.Rows.Count > 0)
                TotalSum(pFoxOrdersDT, ref DestinationDT, "Фокс П-042-4", "Фокс П-042-4",
                    Convert.ToInt32(FrontMargins.pFoxWidth), Convert.ToInt32(FrontMinSizes.pFoxMinWidth), Convert.ToInt32(FrontMargins.pFoxHeight));
            if (pFlorencOrdersDT.Rows.Count > 0)
                TotalSum(pFlorencOrdersDT, ref DestinationDT, "П-1418", "П-01418",
                    Convert.ToInt32(FrontMargins.pFlorencWidth), Convert.ToInt32(FrontMinSizes.pFlorencMinWidth), Convert.ToInt32(FrontMargins.pFlorencHeight));
            if (Techno5OrdersDT.Rows.Count > 0)
            {
                TotalSum(Techno5OrdersDT, ref DestinationDT, "Техно-5 П-516", "Техно-5 П-506",
                    Convert.ToInt32(FrontMargins.Techno5Width), Convert.ToInt32(FrontMinSizes.Techno5MinWidth), Convert.ToInt32(FrontMargins.Techno5Height));
            }

            if (PR1OrdersDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = PR1OrdersDT.Copy();
                for (int i = 0; i < PR1OrdersDT.Rows.Count; i++)
                {
                    object x1 = PR1OrdersDT.Rows[i]["Height"];
                    object x2 = PR1OrdersDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                TotalSum(TempPR1OrdersDT, ref DestinationDT, "Марсель-3 П-141", "Марсель-3 П-041",
                    Convert.ToInt32(FrontMargins.Marsel3Width), Convert.ToInt32(FrontMinSizes.Marsel3MinWidth), Convert.ToInt32(FrontMargins.Marsel3Height));
            }
            if (ShervudOrdersDT.Rows.Count > 0)
            {
                DataTable TempPR1OrdersDT = ShervudOrdersDT.Copy();
                for (int i = 0; i < ShervudOrdersDT.Rows.Count; i++)
                {
                    object x1 = ShervudOrdersDT.Rows[i]["Height"];
                    object x2 = ShervudOrdersDT.Rows[i]["Width"];
                    TempPR1OrdersDT.Rows[i]["Width"] = x1;
                    TempPR1OrdersDT.Rows[i]["Height"] = x2;
                }
                TotalSum(TempPR1OrdersDT, ref DestinationDT, "Шервуд П-042-4", "Шервуд П-043-4",
                    Convert.ToInt32(FrontMargins.ShervudWidth), Convert.ToInt32(FrontMinSizes.ShervudMinWidth), Convert.ToInt32(FrontMargins.ShervudHeight));
            }
            if (PR3OrdersDT.Rows.Count > 0)
            {
                TotalSum(PR3OrdersDT, ref DestinationDT, "Техно-2 П-216", "Техно-2 П-206",
                    Convert.ToInt32(FrontMargins.Techno2Width), Convert.ToInt32(FrontMinSizes.PR3MinWidth), Convert.ToInt32(FrontMargins.Techno2Height));
            }
            if (PRU8OrdersDT.Rows.Count > 0)
            {
                TotalSum(PRU8OrdersDT, ref DestinationDT, "Техно-1 П-116", "Техно-1 П-106",
                    Convert.ToInt32(FrontMargins.Techno1Width), Convert.ToInt32(FrontMinSizes.Techno1MinWidth), Convert.ToInt32(FrontMargins.Techno1Height));
            }
            //DT.Dispose();

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Front, Color";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        #endregion Methods
    }
}