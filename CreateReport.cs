//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;

//using OfficeOpenXml;
//using System.Drawing;

//namespace pagibigEODDataPusher
//{
//    class CreateReport:CSBaseDALMS
//    {

//        public  DataTable GetReport()
//        {
//            DataTable dt = new DataTable();

//            dt = GetDatatable("ReportPagIBIG", CommandType.StoredProcedure);
//            return dt;
//        }

//        public DataTable GetReportv2()
//        {
//            DataTable dt = new DataTable();

//            dt = GetDatatable("ReportPagIBIGv2", CommandType.StoredProcedure);
//            dt = GetDatatable("ReportPagIBIGv3", CommandType.StoredProcedure);
//            return dt;
//        }

//        public void GenerateReport()
//        {
//            DataTable GRDatatable = new DataTable();
//            int totalTrans = 0;

//            string dailyReportRepo = Path.Combine(Application.StartupPath, "DailyReport");
//            string processedRepo = Path.Combine(Application.StartupPath, "Processed");
//            if (!Directory.Exists(dailyReportRepo)) Directory.CreateDirectory(dailyReportRepo);
//            if (!Directory.Exists(processedRepo)) Directory.CreateDirectory(processedRepo);

//            string filename = "DCS_REPORT_" + Convert.ToDateTime(DateTime.Now).ToString("yyyyMMdd") + ".xlsx";
//            if (File.Exists(Path.Combine(dailyReportRepo, filename)))
//            {

//                string source = "";
//                string dest = "";
//                source = Path.Combine(dailyReportRepo, filename);
//                dest = Path.Combine(processedRepo, filename);
//                if (File.Exists(dest)) File.Delete(dest);
//                try
//                {
//                    System.IO.File.Move(source, dest);
//                }
//                catch (IOException e)
//                {
//                    Console.Write("Check if file is open");
//                    return;
//                }
//            }

//            GRDatatable = GetReport();

//            FileInfo newFile = new FileInfo(System.IO.Path.Combine(Application.StartupPath, "DailyReport\\" + filename));
//            using (ExcelPackage xlPck = new ExcelPackage(newFile))
//            {
//                ExcelWorksheet ws = xlPck.Workbook.Worksheets.Add(DateTime.Now.ToString("yyyyMMddHHmmss"));

//                DataTable unq = new DataTable();
//                unq = GRDatatable.DefaultView.ToTable(true, "Branch");

//                ExcelRange rng = ws.Cells["A1:C1"];
//                rng.Merge = true;
//                rng.Style.WrapText = true;                
//                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
//                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
//                rng.Value = "RELEASED CARDS ";
//                rng.Style.Font.Size = 14;
//                rng.Style.Font.Bold = true;
//                rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;                
//                rng.Style.Fill.BackgroundColor.SetColor(Color.Black);
//                rng.Style.Font.Color.SetColor(Color.White);
//                rng.AutoFitColumns();      

//                ws.Column(1).Width = 40;
//                ws.Column(2).Width = 15;
//                ws.Column(3).Width = 15;

//                int intRow = 2;
//                foreach (DataRow rw in unq.Rows)
//                {
//                    int _r = 0;
//                    decimal a = 0;
//                    string rowval = rw[0].ToString();
//                    if (rowval == "")
//                    {
//                        break;
//                    }

//                    DataView LoadFilter = new DataView(GRDatatable);
//                    string filterExp = "Branch ='" + rw[0].ToString() + "'";
//                    LoadFilter.RowFilter = filterExp;

//                    DataTable LF = new DataTable();
//                    LF = LoadFilter.ToTable();

//                    string cond = "Branch="+rw[0].ToString();
//                    totalTrans = Convert.ToInt32(GRDatatable.Compute("Sum(CardCount)", filterExp));                    

//                    headers
//                    PopulateCell(ref ws, "PAGIBIG BRANCH", 1, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);                    
//                    PopulateCell(ref ws, "DATE", 1, 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);                    
//                    PopulateCell(ref ws, "ISSUED", 1, 3, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                    rng.Style.Font.Color.SetColor(Color.White);

//                    int intCol = 0;
//                    int intSumIssuedCards = 0;

//                    foreach (DataRow rwLF in LF.Rows)
//                    {
//                        intCol = 0;

//                        PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10,false);
//                        intCol += 1;
//                        PopulateCell(ref ws, Convert.ToDateTime(rwLF[intCol]).ToString("MM/dd/yyyy"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10,false);
//                        intCol += 1;
//                        PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                        intSumIssuedCards += (int)rwLF[intCol];
//                        ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";

//                        intRow += 1;
//                    }

//                    intCol = 0;
//                    PopulateCell(ref ws, "", intRow, intCol+1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                    intCol += 1;
//                    PopulateCell(ref ws, "TOTAL", intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                    ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
//                    intCol += 1;
//                    PopulateCell(ref ws, intSumIssuedCards.ToString("N0"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                    ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
//                    ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
//                    intRow += 1;

//                    InsertEmpRows(ref ws, intRow, intCol + 1);
//                    intRow += 1;
//                }

//                xlPck.Save();

//                SendMail sendMail = new SendMail();
//                string errMsg = "";
//                sendMail.SendNotification("[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]", "TEST", newFile.FullName, ref errMsg);
//            }
//        }

//        public void GenerateReportv2()
//        {
//            try
//            {
//                WriteToLog("Start of process");

//                DataTable GRDatatable = new DataTable();
//                int totalTrans = 0;

//                string dailyReportRepo = Path.Combine(Application.StartupPath, "DailyReport");
//                string processedRepo = Path.Combine(Application.StartupPath, "Processed");
//                if (!Directory.Exists(dailyReportRepo)) Directory.CreateDirectory(dailyReportRepo);
//                if (!Directory.Exists(processedRepo)) Directory.CreateDirectory(processedRepo);

//                string filename = "DCS_REPORT_" + Convert.ToDateTime(DateTime.Now).ToString("yyyyMMdd") + ".xlsx";
//                if (File.Exists(Path.Combine(dailyReportRepo, filename)))
//                {

//                    string source = "";
//                    string dest = "";
//                    source = Path.Combine(dailyReportRepo, filename);
//                    dest = Path.Combine(processedRepo, filename);
//                    if (File.Exists(dest)) File.Delete(dest);
//                    try
//                    {
//                        System.IO.File.Move(source, dest);
//                    }
//                    catch (IOException e)
//                    {
//                        Console.Write("Check if file is open");
//                        return;
//                    }
//                }

//                GRDatatable = GetReportv2();

//                FileInfo newFile = new FileInfo(System.IO.Path.Combine(Application.StartupPath, "DailyReport\\" + filename));
//                using (ExcelPackage xlPck = new ExcelPackage(newFile))
//                {
//                    ExcelWorksheet ws = xlPck.Workbook.Worksheets.Add(DateTime.Now.ToString("yyyyMMddHHmmss"));

//                    DataTable unq = new DataTable();
//                    unq = GRDatatable.DefaultView.ToTable(true, "Branch");

//                    ExcelRange rng = ws.Cells["A1:F1"];
//                    rng.Merge = true;
//                    rng.Style.WrapText = true;                
//                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
//                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
//                    rng.Value = "RELEASED CARDS ";
//                    rng.Style.Font.Size = 14;
//                    rng.Style.Font.Bold = true;
//                    rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
//                    rng.Style.Fill.BackgroundColor.SetColor(Color.Black);
//                    rng.Style.Font.Color.SetColor(Color.White);
//                    rng.AutoFitColumns();      

//                    ws.Column(1).Width = 40;
//                    ws.Column(2).Width = 15;
//                    ws.Column(3).Width = 15;
//                    ws.Column(4).Width = 15;
//                    ws.Column(5).Width = 15;
//                    ws.Column(6).Width = 15;

//                    int intRow = 2;
//                    foreach (DataRow rw in unq.Rows)
//                    {
//                        int _r = 0;
//                        decimal a = 0;
//                        string rowval = rw[0].ToString();
//                        if (rowval == "")
//                        {
//                            break;
//                        }

//                        DataView LoadFilter = new DataView(GRDatatable);
//                        string filterExp = "Branch ='" + rw[0].ToString() + "'";
//                        LoadFilter.RowFilter = filterExp;

//                        DataTable LF = new DataTable();
//                        LF = LoadFilter.ToTable();

//                        string cond = "Branch="+rw[0].ToString();
//                        totalTrans = Convert.ToInt32(GRDatatable.Compute("Sum(CardCount)", filterExp));                    

//                        headers
//                        PopulateCell(ref ws, "PAG-IBIG BRANCH", 1, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        PopulateCell(ref ws, "DATE", 1, 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        PopulateCell(ref ws, "DELIVERED", 1, 3, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        PopulateCell(ref ws, "ISSUED", 1, 4, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        PopulateCell(ref ws, "SPOILED", 1, 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        PopulateCell(ref ws, "BALANCE", 1, 6, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        rng.Style.Font.Color.SetColor(Color.White);

//                        int intCol = 0;
//                        int intSumDeliveredCards = 0;
//                        int intSumIssuedCards = 0;
//                        int intSumSpoiledCards = 0;

//                        foreach (DataRow rwLF in LF.Rows)
//                        {
//                            intCol = 0;

//                            PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                            intCol += 1;
//                            PopulateCell(ref ws, Convert.ToDateTime(rwLF[intCol]).ToString("MM/dd/yyyy"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                            intCol += 1;
//                            PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                            intSumDeliveredCards += (int)rwLF[intCol];
//                            ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
//                            intCol += 1;
//                            PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                            intSumIssuedCards += (int)rwLF[intCol];
//                            ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
//                            intCol += 1;
//                            PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                            intSumSpoiledCards += (int)rwLF[intCol];
//                            ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
//                            intCol += 1;
//                            PopulateCell(ref ws, rwLF[intCol].ToString(), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                            ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";

//                            intRow += 1;
//                        }

//                        intCol = 0;
//                        PopulateCell(ref ws, "", intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Left, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//                        intCol += 1;
//                        PopulateCell(ref ws, "TOTAL", intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
//                        intCol += 1;
//                        PopulateCell(ref ws, intSumDeliveredCards.ToString("N0"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
//                        ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
//                        intCol += 1;
//                        PopulateCell(ref ws, intSumIssuedCards.ToString("N0"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
//                        ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
//                        intCol += 1;
//                        PopulateCell(ref ws, intSumSpoiledCards.ToString("N0"), intRow, intCol + 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, true);
//                        ws.Cells[intRow, intCol + 1].Style.Numberformat.Format = "#,##0";
//                        ws.Cells[intRow, intCol + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

//                        intRow += 1;

//                        InsertEmpRowsv2(ref ws, intRow, intCol + 1);
//                        intRow += 1;
//                    }

//                    xlPck.Save();

//                    SendMail sendMail = new SendMail();
//                    string errMsg = "";
//                    if (sendMail.SendNotification("[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]", "TEST", newFile.FullName, ref errMsg))
//                    {
//                        WriteToLog("Report successfully sent");
//                    }
//                    else
//                    {
//                        WriteToLog("Failed to send report. Error " + errMsg);
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                WriteToLog("Failed to generate report. Runtime error " + ex.Message);
//            }
//            finally {
//                WriteToLog("End of process");
//            }
//        }

//        private void WriteToLog(string data)
//        {
//            using (StreamWriter sr = new StreamWriter("Log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt", true))
//            {
//                sr.WriteLine(DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt ") + data);
//                sr.Dispose();
//                sr.Close();
//            }
//        }

//        public void PopulateCell(ref ExcelWorksheet ws, string value, int intRow, int intCol, 
//            OfficeOpenXml.Style.ExcelHorizontalAlignment ExcelHorizontalAlignment,
//            OfficeOpenXml.Style.ExcelVerticalAlignment ExcelVerticalAlignment,
//            int ExcelFontSize,           
//            bool IsBold)
//        {
//            ws.Cells[intRow,intCol].Value = value;
//            ws.Cells[intRow, intCol].Style.HorizontalAlignment = ExcelHorizontalAlignment;
//            ws.Cells[intRow, intCol].Style.VerticalAlignment = ExcelVerticalAlignment;
//             ws.Cells["A1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
//             ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.White);
//            ws.Cells[intRow, intCol].Style.Font.Bold = IsBold;
//            ws.Cells[intRow, intCol].Style.Font.Size = ExcelFontSize;
//            ws.Cells[intRow, intCol].Style.Font.Color.SetColor(Color.Black);            
//        }

//        private void InsertEmpRows(ref ExcelWorksheet ws, int intRow, int intCol)
//        {
//            fillers
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//        }

//        private void InsertEmpRowsv2(ref ExcelWorksheet ws, int intRow, int intCol)
//        {
//            fillers
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//            PopulateCell(ref ws, "", intRow, intCol, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment.Top, 10, false);
//        }


//    }

//}
