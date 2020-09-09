using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.ServiceModel.Syndication;
using System.Data;
using NLog.Targets;

namespace pagibigEODDataPusher
{
    class Program
    {
        public static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        private static string reportDateFile = AppDomain.CurrentDomain.BaseDirectory + "reportDate";
        private static string configFile = AppDomain.CurrentDomain.BaseDirectory + "config";
        public static Config config;

        private static DAL dalLocal = null;
        private static DAL dalSys = null;

        enum Process : short
        {
            EOD = 1,
            EmailReport
        }

        public enum bankID
        {
            UBP = 1,
            AUB
        }

        public enum workplaceId
        {
            Onsite = 1,
            Deployment
        }

        public enum eodStatusType
        {
            Pending=1,
            ForApproval,
            ForApprovalRework,
            Disapproved,
            Approved
        }

        public enum consumableId
        {
            Ribbon=1,
            OfficialReceipt,
            CongratulatoryLetter,
            InternetDataLoad,
            CardNew,
            CardRecard
        }   

        //static void Main(string[] args)        
        static void Main()
        {
            logger.Info("Application started");

            //validatations
            Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Initializing...");
            if (!Init())
            {
                CloseConnections();
                logger.Info("Application closed");
                System.Threading.Thread.Sleep(5000);
                Environment.Exit(0);
                return;
            }

            ////string[] args = { "2", "5" };
            ////validate arguments
            //if (args == null)
            //{
            //    logger.Error("Args is null");
            //    Environment.Exit(0);
            //    return;
            //}
            ////else
            ////{
            ////    if (args.Length != 2)
            ////    {
            ////        logger.Error("Args is invalid");
            ////        Environment.Exit(0);
            ////        return;
            ////    }
            ////}

            //if (Convert.ToInt16(args[0]) == (short)Process.EOD) ProcessEODData();
            //else if (Convert.ToInt16(args[0]) == (short)Process.EmailReport) EmailReport();

            //ProcessEODData();
            EmailReport();

            CloseConnections();

            logger.Info("Application closed");
        }

        private static bool Init()
        {         
            try
            {
                //check if another instance is running
                if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1)
                {
                    //Console.WriteLine("Another instrance is running. Application will be closed.");
                    logger.Error("Another instrance is running. Application will be closed.");
                    return false;
                }

                //check if file exists
                if (!File.Exists(configFile))
                {
                    //Console.WriteLine("Config file is missing");
                    logger.Error("Config file is missing");
                    return false;
                }

                try
                {
                    config = new Config();
                    var configData = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Config>>(File.ReadAllText(configFile));
                    config = configData[0];
                    //dal.ConStr = config.DbaseConStr;
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("Error reading config file. Runtime catched error " + ex.Message);
                    logger.Error("Error reading config file. Runtime catched error " + ex.Message);
                    return false;
                }

                if (config.BankID == (short)bankID.UBP) dalLocal = new DAL(config.DbaseConStrUbp);
                else dalLocal = new DAL(config.DbaseConStrAub);
                
                dalSys = new DAL(config.DbaseConStrSys);

                //check dbase connection                
                if (!dalLocal.IsConnectionOK())
                {
                    //Console.WriteLine("Connection to local database failed. " + dalLocal.ErrorMessage);
                    logger.Error("Connection to local database failed. " + dalLocal.ErrorMessage);
                    return false;
                }                

                //check dbase connection                
                if (!dalLocal.IsConnectionOK())
                {
                    //Console.WriteLine("Connection to sys database failed. " + dalSys.ErrorMessage);
                    logger.Error("Connection to sys database failed. " + dalSys.ErrorMessage);
                    return false;
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Runtime catched error " + ex.Message);
                logger.Error("Runtime catched error " + ex.Message);
                return false;
            }

            return true;
        }

        private static void CloseConnections()
        {
            if (dalLocal != null)
            {
                dalLocal.Dispose();
                dalLocal = null;
            }
            if (dalSys != null)
            {
                dalSys.Dispose();
                dalSys = null;
            }
        }

        class EODData
        {
            public string reqBranch { get; set; }
            public string Branch { get; set; }
            public int WorkplaceId { get; set; }
            public int totalCount { get; set; }
            public int nw { get; set; }
            public int ww { get; set; }
        }

        private static string reportDate = "";
        private static string dateToday = DateTime.Now.ToString("yyyy-MM-dd");

        private static bool ProcessEODData()
        {
            EOD eod = null;

            //string reportDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            //string reportDate = "";
            //string dateToday = DateTime.Now.ToString("yyyy-MM-dd");


            ////tempo
            //reportDate = "2020-09-0";
            //dateToday = Convert.ToDateTime(reportDate).AddDays(1).ToString("yyyy-MM-dd");
            //eod = new EOD(reportDate, dateToday);
            //if (!eod.GenerateEndOfDay()) logger.Error("Failed to generate end of day report for " + reportDate);
            //return false;
            ////tempo

            if (File.Exists(reportDateFile))
            {
                reportDate = System.IO.File.ReadAllText(reportDateFile);
                dateToday = Convert.ToDateTime(reportDate).Date.AddDays(1).ToString("yyyy-MM-dd");

                eod = new EOD(reportDate, dateToday);
                if (!eod.GenerateEndOfDay()) logger.Error("Failed to generated end of day report for " + reportDate);
                else
                {
                    eod = null;
                    eod = new EOD(dateToday, dateToday);
                    if (!eod.GenerateEndOfDay()) logger.Error("Failed to generated end of day report for " + dateToday);
                }
            }
            else
            {
                DataTable dtLastTwoEntryDates = null;

                if (!dalLocal.SelectLastTwoEntryDates())
                {
                    logger.Error("Failed to get last 2 dates of member table");
                    return false;
                }
                else dtLastTwoEntryDates = dalLocal.TableResult;

                if (Convert.ToDateTime(dtLastTwoEntryDates.Rows[0][0].ToString()).Date != DateTime.Now.Date)
                {

                    reportDate = Convert.ToDateTime(dtLastTwoEntryDates.Rows[0][0]).ToString("yyyy-MM-dd");
                    eod = new EOD(reportDate, dateToday);
                    if (!eod.GenerateEndOfDay()) logger.Error("Failed to generate end of day report for " + reportDate);
                    return false;
                }
                else
                {
                    reportDate = Convert.ToDateTime(dtLastTwoEntryDates.Rows[1][0]).ToString("yyyy-MM-dd");
                    eod = new EOD(reportDate, dateToday);
                    if (!eod.GenerateEndOfDay()) logger.Error("Failed to generated end of day report for " + reportDate);
                    else
                    {
                        eod = null;
                        eod = new EOD(dateToday, dateToday);
                        if (!eod.GenerateEndOfDay()) logger.Error("Failed to generated end of day report for " + dateToday);
                    }
                }
            }

            return true;
        }

        private static bool EmailReport()
        {
            short oldBankId = config.BankID;

            if (File.Exists(reportDateFile))
            {
                reportDate = System.IO.File.ReadAllText(reportDateFile);
                dateToday = Convert.ToDateTime(reportDate).Date.AddDays(1).ToString("yyyy-MM-dd");
            }
            else
            {
                DataTable dtLastTwoEntryDates = null;

                if (!dalLocal.SelectLastTwoEntryDates())
                {
                    logger.Error("Failed to get last 2 dates of member table");
                    return false;
                }
                else dtLastTwoEntryDates = dalLocal.TableResult;

                if (Convert.ToDateTime(dtLastTwoEntryDates.Rows[0][0].ToString()).Date != DateTime.Now.Date) reportDate = Convert.ToDateTime(dtLastTwoEntryDates.Rows[0][0]).ToString("yyyy-MM-dd");
                else reportDate = Convert.ToDateTime(dtLastTwoEntryDates.Rows[1][0]).ToString("yyyy-MM-dd");
            }

            Reports report = null;

            report = new Reports(reportDate, dateToday);
            string outputFile1 = "";
            string outputFile2 = "";
            string outputFile3 = "";
            string htmlBody1 = "";
            string htmlBody2 = "";
            string htmlBody3 = "";

            DataTable dt1 = null;
            DataTable dt2 = null;

            report.GenerateReport(ref outputFile1, ref htmlBody1, ref dt1);
            if (oldBankId == (short)bankID.UBP) config.BankID = (short)bankID.AUB;
            else config.BankID = (short)bankID.UBP;

            report = null;
            report = new Reports(reportDate, dateToday);
            report.GenerateReport(ref outputFile2, ref htmlBody2, ref dt2);

            report.GenerateExcel(dt1, dt2, ref outputFile3, ref htmlBody3);

            config.BankID = oldBankId;

            SendMail sendMail = new SendMail();
            string errMsg = "";            
            if (sendMail.SendNotification(Program.config, htmlBody3 + "<br><br>" + htmlBody1 + "<br><br>" + htmlBody2, string.Format("Pag-Ibig Daily Monitoring Report - {0}", DateTime.Now.ToString("MM/dd/yyyy")), outputFile1, outputFile2, ref errMsg))
                Program.logger.Info("Report successfully sent");
            else Program.logger.Error("Failed to send report. Error " + errMsg);            

            return true;
        }

        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        private void TempCode()
        {
            //string sourceFile = @"D:\WORK\BAYAMBANG\dd.txt";
            //string source = @"D:\WORK\BAYAMBANG\uploaded at plant conso\4";
            //string desti = @"D:\WORK\BAYAMBANG\test";
            //foreach (string line in File.ReadAllLines(sourceFile))
            //{
            //    if (line.Trim() != "")
            //    {
            //        if (Directory.Exists(source + "\\" + line.Trim()))
            //        {
            //            DirectoryCopy(source + "\\" + line.Trim(), desti + "\\" + line.Trim(),false);
            //        }
            //    }
            //}

            //return;
        }


    }
}
