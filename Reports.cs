using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Windows.Forms;

using OfficeOpenXml;
using System.Drawing;
using NLog;
using System.Security.Cryptography;

namespace pagibigEODDataPusher
{
    class Reports
    {

        private string reportDate;
        private string dateToday;

        public Reports(string reportDate, string dateToday)
        {
            this.reportDate = reportDate;
            this.dateToday = dateToday;
        }

        enum ReportElement
        {
            CardReceived = 1,
            CardIssued = 2,
            CardOnsiteWWarranty = 3,
            CardOnsiteNWarranty = 4,
            CardDeployedWWarranty = 5,
            CardDeployedNWarranty = 6,
            CardSpoiled = 7,
            CardMagError = 8,
            CardBalance = 9,
            CashExpected = 10,
            CashDeposit = 11,
            CashOnsiteByDSA = 12,
            CashOnsiteByBank = 13,
            CashDeployedByDSA = 14,
            CashDeployedByBank = 15,
            CashVariance = 16,
            ConsumablesUsedRibbon = 17,
            ConsumablesUsedOR = 18,
            ConsumablesUsedCL = 19
        }

        private DataTable dtData = null;
        private DataTable dtDataPrev = null;
        private DataTable dtConsumablesInData = null;
        private DataTable dtConsumablesOutData = null;
        private DataTable dtConsumablesInOutData = null;
        private DataTable dtConsumablesOutDataPrev = null;
        private DataTable dtReport = null;
        private DataTable dtReport2 = null;

        public void GetReportData()
        {
            //string startDate = string.Format("{0}-01-01", DateTime.Now.Year.ToString());
            string startDate = string.Format("{0}-01-01", DateTime.Now.Year.ToString());
            string endDate = string.Format("{0}-12-31", DateTime.Now.Year.ToString());

            string mtdStartDate = string.Format("{0}-{1}-01", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString().PadLeft(2, '0'));
            string prevDate = string.Format("{0}-12-31", DateTime.Now.Year.ToString());

            DAL dal = new DAL(Program.config.DbaseConStrSys);

            if (!dal.SelectDailyMonitoringReport(Program.config.BankID.ToString(), startDate, dateToday))
            {
                Program.logger.Error("Failed to get data in SelectDailyMonitoringReport. Error " + dal.ErrorMessage);
            }
            else dtData = dal.TableResult;

            //if (!dal.SelectDailyMonitoringReport(Program.config.BankID.ToString(), mtdStartDate, Convert.ToDateTime(reportDate).AddDays(-1).ToString("yyyy-MM-dd")))
                if (!dal.SelectDailyMonitoringReport(Program.config.BankID.ToString(), mtdStartDate, Convert.ToDateTime(reportDate).AddDays(-1).ToString("yyyy-MM-dd")))
            {
                Program.logger.Error("Failed to get data in SelectDailyMonitoringReport. Error " + dal.ErrorMessage);
            }
            else dtDataPrev = dal.TableResult;

            //if (!dal.SelectConsumables(Program.config.BankID.ToString(), startDate, dateToday, "1"))
            //{
            //    Program.logger.Error("Failed to get data in Consumables In. Error " + dal.ErrorMessage);
            //}
            //else dtConsumablesInData = dal.TableResult;

            if (!dal.SelectConsumables(Program.config.BankID.ToString(), startDate, dateToday, "2"))
            {
                Program.logger.Error("Failed to get data in Consumables In. Error " + dal.ErrorMessage);
            }
            else dtConsumablesOutData = dal.TableResult;

            if (!dal.SelectConsumables(Program.config.BankID.ToString(), mtdStartDate, Convert.ToDateTime(reportDate).AddDays(-1).ToString("yyyy-MM-dd"), "2"))
            {
                Program.logger.Error("Failed to get data in Consumables In. Error " + dal.ErrorMessage);
            }
            else dtConsumablesOutDataPrev = dal.TableResult;

            if (!dal.SelectDailyMonitoringReport2(Program.config.BankID.ToString(), dateToday, dateToday))
            {
                Program.logger.Error("Failed to get data in DailyMonitoringReport2. Error " + dal.ErrorMessage);
            }
            else dtReport2 = dal.TableResult;

            if (!dal.SelectConsumablesInOut(Program.config.BankID.ToString(), dateToday, dateToday))
            {
                Program.logger.Error("Failed to get data in ConsumablesInOut. Error " + dal.ErrorMessage);
            }
            else dtConsumablesInOutData = dal.TableResult;

            dal.Dispose();
            dal = null;
        }

        private void AddElementRow(string reportField, string dbField, string dbType, ref short code)
        {
            DataRow rw = dtReport.NewRow();
            rw[0] = code;
            rw[1] = reportField;
            rw[2] = dbField;
            rw[3] = dbType;
            dtReport.Rows.Add(rw);
            code += 1;
        }

        private void CreateReportTable()
        {
            if (dtReport == null)
            {
                dtReport = new DataTable();
                dtReport.Columns.Add("Code", Type.GetType("System.Int16"));
                dtReport.Columns.Add("ReportField", Type.GetType("System.String"));
                dtReport.Columns.Add("DBField", Type.GetType("System.String"));
                dtReport.Columns.Add("DBType", Type.GetType("System.String"));
                dtReport.Columns.Add("Prev", Type.GetType("System.String"));
                dtReport.Columns.Add("Mtd", Type.GetType("System.String"));
                dtReport.Columns.Add("1", Type.GetType("System.String"));
                dtReport.Columns.Add("2", Type.GetType("System.String"));
                dtReport.Columns.Add("3", Type.GetType("System.String"));
                dtReport.Columns.Add("4", Type.GetType("System.String"));
                dtReport.Columns.Add("5", Type.GetType("System.String"));
                dtReport.Columns.Add("6", Type.GetType("System.String"));
                dtReport.Columns.Add("7", Type.GetType("System.String"));
                dtReport.Columns.Add("8", Type.GetType("System.String"));
                dtReport.Columns.Add("9", Type.GetType("System.String"));
                dtReport.Columns.Add("10", Type.GetType("System.String"));
                dtReport.Columns.Add("11", Type.GetType("System.String"));
                dtReport.Columns.Add("12", Type.GetType("System.String"));
                dtReport.Columns.Add("Yearly", Type.GetType("System.String"));
                dtReport.Columns.Add("Average", Type.GetType("System.String"));

                short code = 1;

                DataRow rw = dtReport.NewRow();

                AddElementRow("Received", "Received_Card", "int", ref code); //1
                AddElementRow("Issued", "Issued_Card", "int", ref code);
                AddElementRow("  W/ Warranty", "WWarranty_Card", "int", ref code);
                AddElementRow("  W/O Warranty", "NWarranty_Card", "int", ref code);
                AddElementRow("  W/ Warranty", "WWarranty_Card", "int", ref code); //5
                AddElementRow("  W/O Warranty", "NWarranty_Card", "int", ref code);
                AddElementRow("Spoiled", "Spoiled_Card", "int", ref code);
                AddElementRow("Magstripe Error", "MagError_Card", "int", ref code); //8
                AddElementRow("Balance (Stocks)", "Balance_Card", "int", ref code);
                AddElementRow("Expected", "Expected_Cash", "dec", ref code);
                AddElementRow("Deposit (Validated)", "Deposited_Cash", "dec", ref code); //11
                AddElementRow("  On-site", "ByDSA_Cash", "dec", ref code);
                AddElementRow("  Deployed", "ByDSA_Cash", "dec", ref code);
                AddElementRow("     By DSA", "ByDSA_Cash", "dec", ref code);
                AddElementRow("     By Bank", "ByBank_Cash", "dec", ref code); //15
                AddElementRow("Variance", "Variance", "dec", ref code);
                AddElementRow("Used Ribbon", "ConsumablesUsedRibbon", "int", ref code);
                AddElementRow("Used Offical Receipt", "ConsumablesUsedOR", "int", ref code);
                AddElementRow("Used Congratulatory Letter", "ConsumablesUsedCL", "int", ref code); //19
            }
            else dtReport.Clear();
        }

        private int GetValueInt(DataTable dtSource, short month, int workplaceId, string field)
        {
            int value = 0;
            if (month > DateTime.Now.Month) return value;
            if (dtSource.Select(string.Format("ReportMonth={0}", month)).Length == 0) return value;

            DataTable dt = dtSource.Select(string.Format("ReportMonth={0}", month)).CopyToDataTable();
            foreach (DataRow rw in dt.Rows)
            {
                if (workplaceId == 0)
                { if (rw[field] != DBNull.Value) value += (int)rw[field]; }
                else
                {
                    if ((int)rw["WorkplaceId"] == workplaceId) if (rw[field] != DBNull.Value) value += (int)rw[field];
                }
            }

            return value;
        }

        private decimal GetValueDecimal(DataTable dtSource, short month, int workplaceId, string field)
        {
            decimal value = 0;
            if (month > DateTime.Now.Month) return value;
            if (dtSource.Select(string.Format("ReportMonth={0}", month)).Length == 0) return value;

            DataTable dt = dtSource.Select(string.Format("ReportMonth={0}", month)).CopyToDataTable();
            foreach (DataRow rw in dt.Rows)
            {
                if (workplaceId == 0)
                { if (rw[field] != DBNull.Value) value += (decimal)rw[field]; }
                else
                { if ((int)rw["WorkplaceId"] == workplaceId) if (rw[field] != DBNull.Value) value += (decimal)rw[field]; }
            }

            return value;
        }

        public void GenerateReport1()
        {
            try
            {
                CreateReportTable();
                GetReportData();

                DataTable dtInt = dtReport.Select("DBType='int'").CopyToDataTable();
                DataTable dtDec = dtReport.Select("DBType='dec'").CopyToDataTable();

                DAL dal = new DAL(Program.config.DbaseConStrSys);

                foreach (DataRow rw in dtInt.Rows)
                {
                    int value = 0;                    
                    short prevMonth = Convert.ToInt16(Convert.ToDateTime(reportDate).Month);

                    switch ((short)rw["Code"])
                    {
                        case 3:
                        case 4:
                            value = GetValueInt(dtDataPrev, prevMonth, (int)Program.workplaceId.Onsite, rw["DBField"].ToString());
                            dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = value.ToString("N0");
                            break;
                        case 5:
                        case 6:
                            value = GetValueInt(dtDataPrev, prevMonth, (int)Program.workplaceId.Deployment, rw["DBField"].ToString());
                            dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = value.ToString("N0");
                            break;
                        case 9:
                            int balanceCard = 0;
                            string startDate = string.Format("{0}-{1}-{2}", DateTime.Now.Year, prevMonth.ToString().PadLeft(2, '0'), "01");
                            string endDate = string.Format("{0}-{1}-{2}", DateTime.Now.Year, prevMonth.ToString().PadLeft(2, '0'), DateTime.DaysInMonth(DateTime.Now.Year, prevMonth));

                            if (dal.SelectEODDepositsDistinctBranchByDateAndBankId(Program.config.BankID.ToString(), startDate, endDate))
                            {
                                DataTable dtDistictBranch = dal.TableResult;

                                foreach (DataRow rwBranch in dtDistictBranch.Rows)
                                {
                                    if (dal.SelectEODDepositsLastBalanceCardByDateAndBankIdAndReBranch(Program.config.BankID.ToString(), rwBranch["requesting_branchcode"].ToString(), endDate))
                                    {
                                        foreach (DataRow rwBalance in dal.TableResult.Rows)
                                        {
                                            balanceCard += (int)rwBalance[0];
                                        }
                                    }
                                }
                            }

                            value = balanceCard;
                            dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = value.ToString("N0");
                            break;
                        case 17:
                            value = GetConsumableBalance(dtConsumablesOutDataPrev, Program.consumableId.Ribbon, prevMonth);
                            dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = value.ToString("N0");
                            break;
                        case 18:
                            value = GetConsumableBalance(dtConsumablesOutDataPrev, Program.consumableId.OfficialReceipt, prevMonth);
                            dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = value.ToString("N0");
                            break;
                        case 19:
                            value = GetConsumableBalance(dtConsumablesOutDataPrev, Program.consumableId.CongratulatoryLetter, prevMonth);
                            dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = value.ToString("N0");
                            break;
                        default:
                            value = GetValueInt(dtDataPrev, prevMonth, 0, rw["DBField"].ToString());
                            dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = value.ToString("N0");
                            break;
                    }


                    //dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = 0;
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Mtd"] = 0;

                    int totalInt = 0;

                    for (short i = 1; i <= 12; i++)
                    {
                        value = 0;

                        switch ((short)rw["Code"])
                        {
                            case 3:
                            case 4:
                                value = GetValueInt(dtData, i, (int)Program.workplaceId.Onsite, rw["DBField"].ToString());
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N0");
                                break;
                            case 5:
                            case 6:
                                value = GetValueInt(dtData, i, (int)Program.workplaceId.Deployment, rw["DBField"].ToString());
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N0");
                                break;
                            case 9:
                                int balanceCard = 0;
                                string startDate = string.Format("{0}-{1}-{2}", DateTime.Now.Year, i.ToString().PadLeft(2, '0'), "01");
                                string endDate = string.Format("{0}-{1}-{2}", DateTime.Now.Year, i.ToString().PadLeft(2, '0'), DateTime.DaysInMonth(DateTime.Now.Year, i));

                                if (dal.SelectEODDepositsDistinctBranchByDateAndBankId(Program.config.BankID.ToString(), startDate, endDate))
                                {
                                    DataTable dtDistictBranch = dal.TableResult;

                                    foreach (DataRow rwBranch in dtDistictBranch.Rows)
                                    {
                                        if (dal.SelectEODDepositsLastBalanceCardByDateAndBankIdAndReBranch(Program.config.BankID.ToString(), rwBranch["requesting_branchcode"].ToString(), endDate))
                                        {
                                            foreach (DataRow rwBalance in dal.TableResult.Rows)
                                            {
                                                balanceCard += (int)rwBalance[0];
                                            }
                                        }
                                    }
                                }

                                value = balanceCard;
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N0");
                                break;
                            case 17:
                                value = GetConsumableBalance(dtConsumablesOutData, Program.consumableId.Ribbon, i);
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N0");
                                break;
                            case 18:
                                value = GetConsumableBalance(dtConsumablesOutData, Program.consumableId.OfficialReceipt, i);
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N0");
                                break;
                            case 19:
                                value = GetConsumableBalance(dtConsumablesOutData, Program.consumableId.CongratulatoryLetter, i);
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N0");
                                break;
                            default:
                                value = GetValueInt(dtData, i, 0, rw["DBField"].ToString());
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N0");
                                break;
                        }

                        //if (i == DateTime.Now.Month) dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Mtd"] = value.ToString("N0");
                        //if (i == 8) dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Mtd"] = value.ToString("N0");

                        totalInt += value;

                    }

                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Yearly"] = totalInt.ToString("N0");
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Average"] = Math.Round(((float)totalInt / DateTime.Now.Month), 2);
                }

                foreach (DataRow rw in dtDec.Rows)
                {
                    decimal totalDec = 0;
                    //decimal decExpectedCash = 0;
                    //decimal decDeposited = 0;

                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = 0;
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Mtd"] = 0;

                    for (short i = 1; i <= 12; i++)
                    {
                        decimal value = 0;

                        switch ((short)rw["Code"])
                        {
                            case 12:
                                value = GetValueDecimal(dtData, i, (int)Program.workplaceId.Onsite, rw["DBField"].ToString());
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N");
                                break;
                            case 13:
                                value = GetValueDecimal(dtData, i, (int)Program.workplaceId.Deployment, rw["DBField"].ToString());
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N");
                                break;
                            default:
                                //if ((short)rw["Code"] != (short)ReportElement.CashVariance) value = GetValueDecimal(dtData, i, 0, rw["DBField"].ToString());
                                //else value = decExpectedCash - decDeposited;

                                value = GetValueDecimal(dtData, i, 0, rw["DBField"].ToString());
                                dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value.ToString("N");

                                //if ((short)rw["Code"] == (short)ReportElement.CashExpected) decExpectedCash = value;
                                //else if ((short)rw["Code"] == (short)ReportElement.CashDeposit) decDeposited = value;

                                break;
                        }

                        //if (i == 8) dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Mtd"] = value.ToString("N");

                        totalDec += value;
                    }

                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Yearly"] = totalDec.ToString("N");
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Average"] = Math.Round((totalDec / DateTime.Now.Month), 2);
                }

                dal.Dispose();
                dal = null;

                foreach (DataRow rwMtd in dtReport.Rows)
                {
                    rwMtd["Mtd"] = rwMtd[(DateTime.Now.Month).ToString()];
                }

                //Pag-ibig Daily Monitoring Report  as of MMMDDYYYY
                //Click here to view the Dashboard in PMS <PMS Dashboard link>

                if (dtData != null)
                {
                    DataTable GRDatatable = new DataTable();

                    int totalTrans = 0;

                    string dailyReportRepo = Path.Combine(Application.StartupPath, "DailyReport");
                    string processedRepo = Path.Combine(Application.StartupPath, "Processed");
                    if (!Directory.Exists(dailyReportRepo)) Directory.CreateDirectory(dailyReportRepo);
                    if (!Directory.Exists(processedRepo)) Directory.CreateDirectory(processedRepo);

                    string filename = "PagibigDailyMonitoringReport_" + Convert.ToDateTime(DateTime.Now).ToString("yyyyMMdd") + ".xlsx";
                    if (File.Exists(Path.Combine(dailyReportRepo, filename)))
                    {

                        string source = "";
                        string dest = "";
                        source = Path.Combine(dailyReportRepo, filename);
                        dest = Path.Combine(processedRepo, filename);
                        if (File.Exists(dest)) File.Delete(dest);
                        try
                        {
                            System.IO.File.Move(source, dest);
                        }
                        catch (IOException e)
                        {
                            Console.Write("Check if file is open");
                            return;
                        }
                    }

                    FileInfo newFile = new FileInfo(System.IO.Path.Combine(Application.StartupPath, "DailyReport\\" + filename));

                    ExcelRange rng = null;

                    using (ExcelPackage xlPck = new ExcelPackage(newFile))
                    {
                        ExcelWorksheet ws = xlPck.Workbook.Worksheets.Add("report1-" + DateTime.Now.ToString("yyyyMMdd"));
                        ws.View.ShowGridLines = false;

                        rng = ws.Cells["B7:O32"];
                        rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                        //DataTable unq = new DataTable();
                        //unq = GRDatatable.DefaultView.ToTable(true, "Branch");

                        //ExcelRange rng = ws.Cells["A1:F1"];
                        //rng.Merge = true;
                        //rng.Style.WrapText = true;                
                        //rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        //rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        //rng.Value = "RELEASED CARDS ";
                        //rng.Style.Font.Size = 14;
                        //rng.Style.Font.Bold = true;
                        //rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                        //rng.Style.Font.Color.SetColor(Color.White);
                        //rng.AutoFitColumns();      

                        short colDividerWidth = 3;
                        ws.Column(1).Width = colDividerWidth;

                        ws.Column(2).Width = 24;//40;

                        ws.Column(3).Width = 10;//15;
                        ws.Column(4).Width = 10;//15;
                        //ws.Column(5).Width = colDividerWidth;

                        //months
                        int monthCount = 5 + (DateTime.Now.Month-1);
                        for (short i = 5; i <= monthCount; i++) ws.Column(i).Width = 10;//15;

                        ws.Column(monthCount + 1).Width = 10;//20;
                        ws.Column(monthCount + 2).Width = 10;//20;

                        int intColBase = 2;
                        int intCol = intColBase;
                        int intRow = 2;

                        if(Program.config.BankID==(short)Program.bankID.UBP) PopulateCell(ref ws, "UNION BANK OF THE PHILIPPINES", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false,"",false);
                        else if (Program.config.BankID == (short)Program.bankID.AUB) PopulateCell(ref ws, "ASIA UNITED BANK", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false, "", false);
                        PopulateCell(ref ws, "PAG-IBIG DAILY MONITORING REPORT AS OF " + DateTime.Now.ToString("MMMM dd, yyyy"), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false, "", false);
                        PopulateCell_Url(ref ws, "Click here to view the Dashboard in PMS " + Program.config.PMSUrl, ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false);
                        //rng.Style.Font.Color.SetColor(Color.White);

                        intRow += 1;

                        //PopulateCell(ref ws, "SUMMARY", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false);
                        //PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false);

                        PopulateHeaders(ref ws, ref intRow, ref intCol);

                        foreach (DataRow rw in dtReport.Rows)
                        {
                            intCol = intColBase;
                            InsertData(ref ws, ref intRow, ref intCol, rw);
                            switch (Convert.ToInt32(rw["Code"]))
                            {
                                case 2:
                                    intCol = intColBase;
                                    PopulateCell(ref ws, "On-site", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false);
                                    break;
                                case 4:
                                    intCol = intColBase;
                                    PopulateCell(ref ws, "Deployed", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false);
                                    break;
                                case 9:
                                    intRow += 1;
                                    intCol = intColBase;
                                    PopulateCell(ref ws, "CASH", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false);
                                    break;
                                case 16:
                                    intRow += 1;
                                    intCol = intColBase;
                                    PopulateCell(ref ws, "CONSUMABLES", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false);
                                    break;
                            }
                        }

                        StringBuilder sbHtml = new StringBuilder();

                        sbHtml.Append("<TABLE style=\"font-size:14px; width: 100%;\">");

                        for (short iRow = 2; iRow <= 32;iRow++)
                        {
                            sbHtml.Append("<TR>");

                            switch (iRow)
                            {
                                case 2:
                                case 3:
                                case 4:
                                case 5:
                                case 18:
                                case 27:
                                    for (short iCol = 2; iCol <= 15; iCol++)
                                    {
                                        if (ws.Cells[iRow, iCol].Value == null) sbHtml.Append(string.Format("<TD colspan=14>{0}</TD>", " "));
                                        else sbHtml.Append(string.Format("<TD colspan=14>{0}</TD>", ws.Cells[iRow, iCol].Value.ToString()));
                                    }                                    
                                    break;
                                default:
                                    for (short iCol = 2; iCol <= 15; iCol++)
                                    {
                                        string textAlignCenter = "";
                                        string textWidth = "";
                                        if (iCol == 2) textWidth = "width: 20%;";
                                        else
                                        {
                                            textAlignCenter = "text-align:center;";
                                            textWidth = "width: 7%;";
                                        }
                                        if (ws.Cells[iRow, iCol].Value == null) sbHtml.Append(string.Format("<TD style=\"border: 1px solid black; border-collapse: collapse; {0}{1}\">{2}</TD>", textAlignCenter, textWidth, " "));
                                        else sbHtml.Append(string.Format("<TD style=\"border: 1px solid black; border-collapse: collapse; {0}{1}\">{2}</TD>", textAlignCenter, textWidth, ws.Cells[iRow, iCol].Value.ToString()));
                                    }
                                    break;

                            }                            
                            sbHtml.Append("</TR>");
                        }

                        sbHtml.Append("</TABLE>");

                        //report 2
                        ws = null;
                        ws = xlPck.Workbook.Worksheets.Add("report2-" + DateTime.Now.ToString("yyyyMMdd"));
                        ws.View.ShowGridLines = false;

                        rng = ws.Cells["B5:D5"];                        
                        rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;                                                
                        rng.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Merge = true;

                        rng = ws.Cells["E5:K5"];                        
                        rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;                        
                        rng.Merge = true;                        
                        //rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        rng = ws.Cells["L5:R5"];                        
                        rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;                        
                        rng.Merge = true;
                        //ws.Cells[rng.Address].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        //rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        rng = ws.Cells["S5:U5"];                        
                        rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;                        
                        rng.Merge = true;
                        //ws.Cells[rng.Address].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        //ExcelRange rng = ws.Cells["B7:O32"];
                        //rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);                        

                        ws.Column(1).Width = colDividerWidth;
                        ws.Column(2).Width = 30;
                        ws.Column(3).Width = 7; //15;
                        ws.Column(4).Width = 12;//15;

                        for (short i = 5; i <= 21; i++) ws.Column(i).Width = 15;                        
                        
                        intCol = intColBase;
                        intRow = 2;

                        if (Program.config.BankID == (short)Program.bankID.UBP) PopulateCell(ref ws, "UNION BANK OF THE PHILIPPINES", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false, "", false);
                        else if (Program.config.BankID == (short)Program.bankID.AUB) PopulateCell(ref ws, "ASIA UNITED BANK", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false, "", false);
                        PopulateCell(ref ws, "PAG-IBIG DAILY MONITORING REPORT AS OF " + DateTime.Now.ToString("MMMM dd, yyyy"), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false, "", false);
                        intRow += 1;

                        intCol = 5;
                        PopulateCell(ref ws, "CARD", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, false, "", false);

                        intCol = 12;
                        PopulateCell(ref ws, "CASH", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, false, "", false);

                        intCol = 19;
                        PopulateCell(ref ws, "CONSUMABLES", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, false, "", false);

                        intCol = intColBase;
                        
                        PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, false);

                        PopulateHeadersReport2(ref ws, ref intRow, ref intCol);                        

                        foreach (DataRow rw in dtReport2.Rows)
                        {
                            intCol = intColBase;

                            rw["RemRibbon"] = GetConsumablesRemainingByBranchAndConsumable(rw["requesting_branchcode"].ToString().Trim(), Program.consumableId.Ribbon);
                            rw["RemOR"] = GetConsumablesRemainingByBranchAndConsumable(rw["requesting_branchcode"].ToString().Trim(), Program.consumableId.OfficialReceipt);
                            rw["RemCL"] = GetConsumablesRemainingByBranchAndConsumable(rw["requesting_branchcode"].ToString().Trim(), Program.consumableId.CongratulatoryLetter);

                            for (short i = 0; i <= dtReport2.Columns.Count - 2; i++)
                            {
                                switch (i)
                                {
                                    case 3:
                                    case 4:
                                    case 5:
                                    case 6:
                                    case 7:
                                    case 8:
                                    case 9:
                                    case 17:
                                    case 18:
                                    case 19:
                                        PopulateCell(ref ws, Convert.ToInt64(rw[i]), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false,true, "#,##0;(#,##0)",true);
                                        break;
                                    case 10:
                                    case 11:
                                    case 12:
                                    case 13:
                                    case 14:
                                        PopulateCell(ref ws, Convert.ToDecimal(rw[i]), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, "#,##0.00;(#,##0.00)", true);
                                        break;
                                    default:                                        
                                        PopulateCell(ref ws, rw[i].ToString(), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, "", true);
                                        break;
                                }    
                                
                            }
                            intRow += 1;
                            //PopulateCell(ref ws, Convert.ToInt64(rw[dtReport2.Columns.Count - 1]), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, true,true, "#,##0;(#,##0)",true);                            
                        }
                        

                        xlPck.Save();

                        SendMail sendMail = new SendMail();
                        string errMsg = "";
                        //if (sendMail.SendNotification(Program.config, "[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]", "TEST", newFile.FullName, ref errMsg))
                        if (sendMail.SendNotification(Program.config, sbHtml.ToString(), "Pag-ibig Daily Monitoring Report - " + DateTime.Now.ToShortDateString(), newFile.FullName, ref errMsg))
                        {
                            Program.logger.Info("Report successfully sent");
                        }
                        else
                        {
                            Program.logger.Error("Failed to send report. Error " + errMsg);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.logger.Error("Failed to generate report. Runtime error " + ex.Message);
            }
            finally
            {
                //WriteToLog("End of process");
            }
        }

        private int GetConsumablesRemainingByBranchAndConsumable(string reqBranch, Program.consumableId consumableId)
        {
            if (dtConsumablesInOutData.Select(string.Format("requesting_branchcode='{0}' AND ConsumableID={1}", reqBranch, (int)consumableId)).Length == 0) return 0;
            else
            {
                int inValue = 0;
                int outValue = 0;
                foreach (DataRow rw in dtConsumablesInOutData.Select(string.Format("requesting_branchcode='{0}' AND ConsumableID={1}", reqBranch, (int)consumableId)))
                {
                    if (rw["TransactionTypeID"].ToString() == "1") inValue = (int)rw["Quantity"];
                    else if (rw["TransactionTypeID"].ToString() == "2") outValue = (int)rw["Quantity"];
                }

                return inValue - outValue;
            }
        }

        private int GetConsumableBalance(DataTable dtSourceOut, Program.consumableId consumableId, int reportMonth)
        {
            //int inValue = 0;
            int outValue = 0;
            //if (dtConsumablesInData.Select(string.Format("ConsumableID={0} AND ReportMonth={1}", (int)consumableId, reportMonth)).Length > 0)inValue = (int)dtConsumablesInData.Select(string.Format("ConsumableID={0} AND ReportMonth={1}", (int)consumableId, reportMonth))[0]["Total"];
            if (dtSourceOut.Select(string.Format("ConsumableID={0} AND ReportMonth={1}", (int)consumableId, reportMonth)).Length > 0) outValue = (int)dtSourceOut.Select(string.Format("ConsumableID={0} AND ReportMonth={1}", (int)consumableId, reportMonth))[0]["Total"];
            //return inValue - outValue;
            return outValue;
        }

        private void PopulateHeaders(ref ExcelWorksheet ws, ref int intRow, ref int intCol)
        {
            PopulateCell(ref ws, "CARD", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "PREV", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "MTD", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            //intCol++;

            for (short i = 1; i <= DateTime.Now.Month; i++)
            {
                string monthAbrev = Convert.ToDateTime(string.Format(@"{0}/01/{1}", i.ToString(), DateTime.Now.Year)).ToString("MMM");

                PopulateCell(ref ws, monthAbrev.ToUpper(), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            }

            PopulateCell(ref ws, "YEAR", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);

            //last row should have isIncrement=true
            PopulateCell(ref ws, "AVERAGE".ToUpper(), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, true);
        }

        private void PopulateHeadersReport2(ref ExcelWorksheet ws, ref int intRow, ref int intCol)
        {
            PopulateCell(ref ws, "Branch", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Bank", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Workplace", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Received", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Issued", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "W/ Warranty", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "W/O Warranty", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Spoiled", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Magstripe Error", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Balance (Stocks)", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Expcected", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Deposited", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "By DSA", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "By Bank", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Variance", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Depository Bank", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Status", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Rem. Ribbon", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);
            PopulateCell(ref ws, "Rem. OR", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, false, true);

            //last row should have isIncrement=true
            PopulateCell(ref ws, "Rem. CL".ToUpper(), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true, true, true);
        }

        private void InsertData(ref ExcelWorksheet ws, ref int intRow, ref int intCol, DataRow rw)
        {
            string ExcelNumberFormat = "";
            if (rw["DBType"].ToString() == "dec") ExcelNumberFormat = "#,##0.00;(#,##0.00)";
            else if (rw["DBType"].ToString() == "int") ExcelNumberFormat = "#,##0;(#,##0)";
            PopulateCell(ref ws, rw["ReportField"].ToString(), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);

            string value = "";

            value = rw["Prev"].ToString().Replace(",", "");
            if (rw["DBType"].ToString() == "dec") PopulateCell(ref ws, Convert.ToDecimal(value), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, ExcelNumberFormat);
            else if (rw["DBType"].ToString() == "int") PopulateCell(ref ws, Convert.ToInt64(value), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, ExcelNumberFormat);

            value = rw["Mtd"].ToString().Replace(",", "");
            if (rw["DBType"].ToString() == "dec") PopulateCell(ref ws, Convert.ToDecimal(value), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, ExcelNumberFormat);
            else if (rw["DBType"].ToString() == "int") PopulateCell(ref ws, Convert.ToInt64(value), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, ExcelNumberFormat);

            //intCol++;

            for (short i = 1; i <= DateTime.Now.Month; i++)
            {
                try
                {
                    value = rw[i.ToString()].ToString().Replace(",", "");
                    if (rw["DBType"].ToString() == "dec") PopulateCell(ref ws, Convert.ToDecimal(value), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, ExcelNumberFormat);
                    else if (rw["DBType"].ToString() == "int") PopulateCell(ref ws, Convert.ToInt64(value), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, ExcelNumberFormat);
                }
                catch (Exception ex)
                {
                    Program.logger.Error(ex.Message);
                }
            }

            value = rw["Yearly"].ToString().Replace(",", "");
            if (rw["DBType"].ToString() == "dec") PopulateCell(ref ws, Convert.ToDecimal(value), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, ExcelNumberFormat);
            else if (rw["DBType"].ToString() == "int") PopulateCell(ref ws, Convert.ToInt64(value), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true, ExcelNumberFormat);

            //last row should have isIncrement=true
            PopulateCell(ref ws, Convert.ToDecimal(rw["Average"]), ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, true, true, "#,##0.00;(#,##0.00)");
        }

        public void PopulateCell(ref ExcelWorksheet ws, object value, ref int intRow, ref int intCol,
           OfficeOpenXml.Style.ExcelHorizontalAlignment ExcelHorizontalAlignment,
           OfficeOpenXml.Style.ExcelVerticalAlignment ExcelVerticalAlignment,
           int ExcelFontSize,
           bool IsBold,
           bool isIncrementRow = false,
           bool isIncrementColumn = true,
           string ExcelNumberFormat = "",
           bool IsPutCellBorders = true)
        {

            ws.Cells[intRow, intCol].Value = value;
            ws.Cells[intRow, intCol].Style.HorizontalAlignment = ExcelHorizontalAlignment;
            ws.Cells[intRow, intCol].Style.VerticalAlignment = ExcelVerticalAlignment;
            //ws.Cells[intRow, intCol].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            // ws.Cells["A1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            // ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.White);
            ws.Cells[intRow, intCol].Style.Font.Bold = IsBold;
            ws.Cells[intRow, intCol].Style.Font.Size = ExcelFontSize;
            ws.Cells[intRow, intCol].Style.Font.Color.SetColor(Color.Black);
            if (ExcelNumberFormat != "") ws.Cells[intRow, intCol].Style.Numberformat.Format = ExcelNumberFormat;
            //ws.Cells[intRow, intCol].Style.Numberformat.Format = "YYYY-MM";
            //ws.Cells[intRow, intCol].Style.Numberformat.Format = "#,##0.0";     

            if (IsPutCellBorders)
            {
                ws.Cells[intRow, intCol].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[intRow, intCol].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[intRow, intCol].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[intRow, intCol].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            if (isIncrementRow) intRow++;
            if (isIncrementColumn) intCol++;
        }

        public void PopulateCell_Url(ref ExcelWorksheet ws, object value, ref int intRow, ref int intCol,
           OfficeOpenXml.Style.ExcelHorizontalAlignment ExcelHorizontalAlignment,
           OfficeOpenXml.Style.ExcelVerticalAlignment ExcelVerticalAlignment,
           int ExcelFontSize,
           bool IsBold,
           bool isIncrementRow = false,
           bool isIncrementColumn = true,
           string ExcelNumberFormat = "")
        {
            ws.Cells[intRow, intCol].Hyperlink = new Uri(Program.config.PMSUrl);
            ws.Cells[intRow, intCol].Value = value;
            ws.Cells[intRow, intCol].Style.HorizontalAlignment = ExcelHorizontalAlignment;
            ws.Cells[intRow, intCol].Style.VerticalAlignment = ExcelVerticalAlignment;
            // ws.Cells["A1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            // ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.White);
            ws.Cells[intRow, intCol].Style.Font.Bold = IsBold;
            ws.Cells[intRow, intCol].Style.Font.Size = ExcelFontSize;
            ws.Cells[intRow, intCol].Style.Font.Color.SetColor(Color.Blue);
            if (ExcelNumberFormat != "") ws.Cells[intRow, intCol].Style.Numberformat.Format = ExcelNumberFormat;
            //ws.Cells[intRow, intCol].Style.Numberformat.Format = "YYYY-MM";
            //ws.Cells[intRow, intCol].Style.Numberformat.Format = "#,##0.0";            

            if (isIncrementRow) intRow++;
            if (isIncrementColumn) intCol++;
        }

        private void InsertEmpRows(ref ExcelWorksheet ws, ref int intRow, ref int intCol)
        {
            //fillers
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true);
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true);
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, true, true);
        }

        private void InsertEmptyRowsv2(ref ExcelWorksheet ws, ref int intRow, ref int intCol)
        {
            //fillers
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true);
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true);
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true);
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true);
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, false, true);
            PopulateCell(ref ws, "", ref intRow, ref intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false, true, true);
        }

    }
}
