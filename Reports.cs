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

namespace pagibigEODDataPusher
{
    class Reports
    {

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


        private Config config;
        private NLog.Logger logger;
        private DataTable dtData = null;
        private DataTable dtReport = null;

        public void GetReportData()
        {            
            DAL dal = new DAL(config.DbaseConStrSys);
            if (!dal.SelectDailyMonitoringReport())
            {
                logger.Error("Failed to get data in SelectDailyMonitoringReport. Error " + dal.ErrorMessage);
            }
            else dtData = dal.TableResult;
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

                //dtReport.Columns.Add("Jan", Type.GetType("System.String"));
                //dtReport.Columns.Add("Feb", Type.GetType("System.String"));
                //dtReport.Columns.Add("Mar", Type.GetType("System.String"));
                //dtReport.Columns.Add("Apr", Type.GetType("System.String"));
                //dtReport.Columns.Add("May", Type.GetType("System.String"));
                //dtReport.Columns.Add("Jun", Type.GetType("System.String"));
                //dtReport.Columns.Add("Jul", Type.GetType("System.String"));
                //dtReport.Columns.Add("Aug", Type.GetType("System.String"));
                //dtReport.Columns.Add("Sep", Type.GetType("System.String"));
                //dtReport.Columns.Add("Oct", Type.GetType("System.String"));
                //dtReport.Columns.Add("Nov", Type.GetType("System.String"));
                //dtReport.Columns.Add("Dec", Type.GetType("System.String"));

                short code = 1;

                DataRow rw = dtReport.NewRow();
                AddElementRow("CardReceived", "Received_Card", "int", ref code);
                AddElementRow("CardIssued", "Issued_Card", "int", ref code);
                AddElementRow("CardOnsiteWWarranty", "WWarranty_Card", "int", ref code);
                AddElementRow("CardOnsiteNWarranty", "NWarranty_Card", "int", ref code);
                AddElementRow("CardDeployedWWarranty", "WWarranty_Card", "int", ref code);
                AddElementRow("CardDeployedNWarranty", "NWarranty_Card", "int", ref code);
                AddElementRow("CardSpoiled", "Spoiled_Card", "int", ref code);
                AddElementRow("CardMagError", "MagError_Card", "int", ref code);
                AddElementRow("CardBalance", "Balance_Card", "int", ref code);
                AddElementRow("CashExpected", "Expected_Cash", "dec", ref code);
                AddElementRow("CashDeposit", "Deposited_Cash", "dec", ref code);
                AddElementRow("CashOnsiteByDSA", "ByDSA_Cash", "dec", ref code);
                AddElementRow("CashOnsiteByBank", "ByBank_Cash", "dec", ref code);
                AddElementRow("CashDeployedByDSA", "ByDSA_Cash", "dec", ref code);
                AddElementRow("CashDeployedByBank", "ByBank_Cash", "dec", ref code);
                AddElementRow("CashVariance", "Variance", "dec", ref code);
                AddElementRow("ConsumablesUsedRibbon", "ConsumablesUsedRibbon", "int", ref code);
                AddElementRow("ConsumablesUsedOR", "ConsumablesUsedOR", "int", ref code);
                AddElementRow("ConsumablesUsedCL", "ConsumablesUsedCL", "int", ref code);
            }
            else dtReport.Clear();
        }

        private int GetValueInt(short month, short workplaceId, string field)
        {
            int value = 0;
            if (month > DateTime.Now.Month) return value;
            if (dtData.Select(string.Format("ReportMonth={0}", month)).Length == 0) return value;

            DataTable dt = dtData.Select(string.Format("ReportMonth={0}", month)).CopyToDataTable();
            foreach (DataRow rw in dt.Rows)
            {
                if (workplaceId == 0)
                { if (rw[field] != DBNull.Value) value += (int)rw[field]; }
                else
                {
                    if ((short)rw["WorkplaceId"] == workplaceId) if (rw[field] != DBNull.Value) value += (int)rw[field];
                }
            }

            return value;
        }

        private decimal GetValueDecimal(short month, short workplaceId, string field)
        {
            decimal value = 0;
            if (month > DateTime.Now.Month) return value;
            if (dtData.Select(string.Format("ReportMonth={0}", month)).Length == 0) return value;

            DataTable dt = dtData.Select(string.Format("ReportMonth={0}", month)).CopyToDataTable();
            foreach (DataRow rw in dt.Rows)
            {
                if (workplaceId == 0)
                { if (rw[field] != DBNull.Value) value += (decimal)rw[field]; }
                else
                { if ((short)rw["WorkplaceId"] == workplaceId) if (rw[field] != DBNull.Value) value += (decimal)rw[field]; }
            }

            return value;
        }

        public void GenerateReportv2(Config config, NLog.Logger logger)
        {
            try
            {
                this.config = config;
                this.logger = logger;

                CreateReportTable();
                GetReportData();

                DataTable dtInt = dtReport.Select("DBType='int'").CopyToDataTable();
                DataTable dtDec = dtReport.Select("DBType='dec'").CopyToDataTable();

                int totalInt = 0;
                decimal totalDec = 0;

                foreach (DataRow rw in dtInt.Rows)
                {
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = 0;
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Mtd"] = 0;

                    for (short i = 1; i <= 12; i++)
                    {
                        int value = GetValueInt(i, 0, rw["DBField"].ToString());
                        dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value;
                        totalInt += value;
                    }

                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Yearly"] = totalInt.ToString("N0");
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Average"] = (totalInt / DateTime.Now.Month);
                }

                foreach (DataRow rw in dtDec.Rows)
                {
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Prev"] = 0;
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Mtd"] = 0;

                    for (short i = 1; i <= 12; i++)
                    {
                        decimal value = GetValueDecimal(i, 0, rw["DBField"].ToString());
                        dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0][i.ToString()] = value;
                        totalDec += value;
                    }

                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Yearly"] = totalDec.ToString("N2");
                    dtReport.Select(string.Format("Code={0}", (short)rw["Code"]))[0]["Average"] = Math.Round((totalDec / DateTime.Now.Month),2);
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
                    using (ExcelPackage xlPck = new ExcelPackage(newFile))
                    {
                        ExcelWorksheet ws = xlPck.Workbook.Worksheets.Add(DateTime.Now.ToString("yyyyMMddHHmmss"));

                        DataTable unq = new DataTable();
                        unq = GRDatatable.DefaultView.ToTable(true, "Branch");

                        ExcelRange rng = ws.Cells["A1:F1"];
                        //rng.Merge = true;
                        //rng.Style.WrapText = true;                
                        //rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        //rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        //rng.Value = "RELEASED CARDS ";
                        rng.Style.Font.Size = 14;
                        rng.Style.Font.Bold = true;
                        rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                        //rng.Style.Font.Color.SetColor(Color.White);
                        //rng.AutoFitColumns();      

                        ws.Column(1).Width = 40;
                        ws.Column(2).Width = 15;
                        ws.Column(3).Width = 15;
                        ws.Column(4).Width = 15;
                        ws.Column(5).Width = 15;
                        ws.Column(6).Width = 15;

                        int intRow = 4;
                        foreach (DataRow rw in unq.Rows)
                        {
                            //int _r = 0;
                            decimal a = 0;
                            string rowval = rw[0].ToString();
                            if (rowval == "")
                            {
                                break;
                            }

                            DataView LoadFilter = new DataView(GRDatatable);
                            string filterExp = "Branch ='" + rw[0].ToString() + "'";
                            LoadFilter.RowFilter = filterExp;

                            DataTable LF = new DataTable();
                            LF = LoadFilter.ToTable();                  

                            //headers
                            PopulateCell(ref ws, "PAG-IBIG DAILY MONITORING REPORT AS OF " + DateTime.Now.ToString("MMMM dd, yyyy"), 1, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
                            PopulateCell(ref ws, "Click here to view the Dashboard in PMS <PMS Dashboard link>", 1, 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);                            
                            rng.Style.Font.Color.SetColor(Color.White);

                            int intCol = 0;
                            //int intSumDeliveredCards = 0;
                            //int intSumIssuedCards = 0;
                            //int intSumSpoiledCards = 0;

                            //foreach (DataRow rwLF in LF.Rows)
                            //{
                            //    intCol = 0;

                            //    PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
                            //    intCol += 1;
                            //    PopulateCell(ref ws, Convert.ToDateTime(rwLF[intCol]).ToString("MM/dd/yyyy"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
                            //    intCol += 1;
                            //    PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
                            //    intSumDeliveredCards += (int)rwLF[intCol];
                            //    ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
                            //    intCol += 1;
                            //    PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
                            //    intSumIssuedCards += (int)rwLF[intCol];
                            //    ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
                            //    intCol += 1;
                            //    PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
                            //    intSumSpoiledCards += (int)rwLF[intCol];
                            //    ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
                            //    intCol += 1;
                            //    PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
                            //    ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";

                            //    intRow += 1;
                            //}

                            //intCol = 0;
                            //PopulateCell(ref ws, "", intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
                            //intCol += 1;
                            //PopulateCell(ref ws, "TOTAL", intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
                            //ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            //intCol += 1;
                            //PopulateCell(ref ws, intSumDeliveredCards.ToString("N0"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
                            //ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
                            //ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            //intCol += 1;
                            //PopulateCell(ref ws, intSumIssuedCards.ToString("N0"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
                            //ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
                            //ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            //intCol += 1;
                            //PopulateCell(ref ws, intSumSpoiledCards.ToString("N0"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
                            //ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
                            //ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                            intRow += 1;

                            InsertEmpRowsv2(ref ws, intRow, intCol + 1);
                            intRow += 1;
                        }

                        xlPck.Save();

                        //SendMail sendMail = new SendMail();
                        //string errMsg = "";
                        //if (sendMail.SendNotification(config, "[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]", "TEST", newFile.FullName, ref errMsg))
                        //{
                        //    //WriteToLog("Report successfully sent");
                        //}
                        //else
                        //{
                        //    //WriteToLog("Failed to send report. Error " + errMsg);
                        //}
                    }
                }                
            }
            catch (Exception ex)
            {
                logger.Error("Failed to generate report. Runtime error " + ex.Message);                
            }
            finally
            {
                //WriteToLog("End of process");
            }
        }

        public void PopulateCell(ref ExcelWorksheet ws, string value, int intRow, int intCol,
           OfficeOpenXml.Style.ExcelHorizontalAlignment ExcelHorizontalAlignment,
           OfficeOpenXml.Style.ExcelVerticalAlignment ExcelVerticalAlignment,
           int ExcelFontSize,
           bool IsBold)
        {
            ws.Cells[intRow, intCol].Value = value;
            ws.Cells[intRow, intCol].Style.HorizontalAlignment = ExcelHorizontalAlignment;
            ws.Cells[intRow, intCol].Style.VerticalAlignment = ExcelVerticalAlignment;
            // ws.Cells["A1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            // ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.White);
            ws.Cells[intRow, intCol].Style.Font.Bold = IsBold;
            ws.Cells[intRow, intCol].Style.Font.Size = ExcelFontSize;
            ws.Cells[intRow, intCol].Style.Font.Color.SetColor(Color.Black);
        }

        private void InsertEmpRows(ref ExcelWorksheet ws, int intRow, int intCol)
        {
            //fillers
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
        }

        private void InsertEmpRowsv2(ref ExcelWorksheet ws, int intRow, int intCol)
        {
            //fillers
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
        }

    }
}
