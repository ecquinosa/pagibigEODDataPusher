﻿using System;
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

            ProcessEODData();
            //EmailReport();

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
                    logger.Error("Another instance is running. Application will be closed.");
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

            if (!dalSys.GenerateConsumbalesDailyEnv(config.BankID.ToString(), dateToday))
            {
                logger.Error("GenerateConsumbalesDailyEnv() failed. Error " + dalSys.ErrorMessage);
                return false;
            }

            return true;
        }

        private static bool EmailReport()
        {
            Reports report = null;            
            string outputFile1 = "";
            string outputFile2 = "";
            string outputFile3 = "";

            //added on 09/16/2020. check first if emailbody already exist. if exist 
            if (File.Exists(Reports.EmailBodyFile(DateTime.Now.ToString("yyyy-MM-dd"))))
            {
                if (DateTime.Now.Date != Convert.ToDateTime(config.LastSuccessEmailSend).Date)
                {
                    foreach (string file in Directory.GetFiles(Reports.GetDailyReportRepo(DateTime.Now.ToString("yyyy-MM-dd"))))
                    {
                        if (Path.GetExtension(file).ToUpper() == ".XLSX")
                        {
                            if (Path.GetFileNameWithoutExtension(file).Contains("UBP")) outputFile1 = file;
                            else if (Path.GetFileNameWithoutExtension(file).Contains("AUB")) outputFile2 = file;
                        }
                    }

                    SendEmail(System.IO.File.ReadAllText(Reports.EmailBodyFile(DateTime.Now.ToString("yyyy-MM-dd"))), outputFile1, outputFile2);

                    return true;
                }
                else return true;
            }
            else
            {
                short oldBankId = config.BankID;

                Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Getting last 2 dates...");
                logger.Info("Getting last 2 dates");
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

                string htmlBody1 = "";
                string htmlBody2 = "";
                string htmlBody3 = "";

                DataTable dt1 = null;
                DataTable dt2 = null;
             
                report = new Reports(reportDate, dateToday);
                Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Generating report for bankId " + config.BankID.ToString() + "...");
                logger.Info("Generating report for bankId " + config.BankID.ToString());
                if (report.GenerateReport(ref outputFile1, ref htmlBody1, ref dt1))
                {
                    if (oldBankId == (short)bankID.UBP) config.BankID = (short)bankID.AUB;
                    else config.BankID = (short)bankID.UBP;

                    report = null;
                    report = new Reports(reportDate, dateToday);
                    Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Generating report for bankId " + config.BankID.ToString() + "...");
                    logger.Info("Generating report for bankId " + config.BankID.ToString());
                    if (report.GenerateReport(ref outputFile2, ref htmlBody2, ref dt2))
                    {
                        Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Generating excel...");
                        logger.Info("Generating excel");
                        report.GenerateExcelv2(dt1, dt2, ref outputFile3, ref htmlBody3);

                        config.BankID = oldBankId;

                        string emailBody = htmlBody3 + "<br><br>" + htmlBody1 + "<br><br>" + htmlBody2;
                        report.SaveDailyEmailBody(emailBody);

                        SendEmail(emailBody, outputFile1, outputFile2);
                        return true;
                    }
                    else
                    {
                        config.BankID = oldBankId;
                        return false;
                    }
                }
                else
                {
                    config.BankID = oldBankId;
                    return false;
                }
            }
        }

        private static void SendEmail(string emailBody, string outputFile1, string outputFile2)
        {
            if (config.IsSendEmail == 1)
            {
                SendMail sendMail = new SendMail();
                try
                {
                    Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Sending email...");
                    logger.Info("Sending email");

                    string errMsg = "";
                    if (sendMail.SendNotification(Program.config, emailBody, string.Format("Pag-Ibig Daily Monitoring Report - {0}", DateTime.Now.ToString("MM/dd/yyyy")), outputFile1, outputFile2, ref errMsg))
                    {
                        logger.Info("Report successfully sent");

                        config.LastSuccessEmailSend = DateTime.Now.ToString("yyyy-MM-dd");
                        var configs = new List<Config>();
                        configs.Add(config);

                        System.IO.File.WriteAllText(configFile, Newtonsoft.Json.JsonConvert.SerializeObject(configs));
                    }
                    else
                    {
                        logger.Error("Failed to send report. Error " + errMsg);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error("Failed to send report. Error " + ex.Message);
                }
                finally
                { sendMail = null; }
            }
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
