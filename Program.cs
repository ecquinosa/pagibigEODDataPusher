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

        private static string reportDateFile = AppDomain.CurrentDomain.BaseDirectory + "reportDate";
        private static string configFile = AppDomain.CurrentDomain.BaseDirectory + "config";
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static Config config;

        private static DAL dalLocal = null;
        private static DAL dalSys = null;

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

        static void Main()
        {
            //string sourceFile = @"D:\ACCPAGIBIGPH3\accpagibigph3srv\AUB Missing KYC\dd.txt";
            //string source = @"D:\ACCPAGIBIGPH3\AUB\PACKUPDATA\DONE\2020-06-29";
            //string desti = @"D:\ACCPAGIBIGPH3\AUB\PACKUPDATA\FOR_TRANSFER\06292020_MISSINGFILES";
            //foreach (string line in File.ReadAllLines(sourceFile))
            //{
            //    if (line.Trim() != "")
            //    {
            //        if (File.Exists(source + "\\" + line.Trim() + ".zip"))
            //        {
            //            File.Copy(source + "\\" + line.Trim() + ".zip", desti + "\\" + line.Trim() + ".zip");
            //        }
            //    }
            //}

            //return;


            //logger.Info("Application started");
            //Console.Write(DateTime.Now.DayOfWeek);
            //Console.ReadLine();

            //return;

            //validatations
            Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Initializing...");
            if (!Init())
            {
                Console.Write("Init error");
                Console.ReadLine();                
                return;
            }
            //else
            //{
            //    Console.Write("init success");
            //    Console.ReadLine(); 
            //}

            ProcessEODData();

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
                    logger.Info("Another instrance is running. Application will be closed.");
                    return false;
                }

                //check if file exists
                if (!File.Exists(configFile))
                {
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
                    logger.Error("Error reading config file. Runtime catched error " + ex.Message);
                    return false;
                }

                if (config.BankID == 1) dalLocal = new DAL(config.DbaseConStrUbp);
                else dalLocal = new DAL(config.DbaseConStrAub);
                
                dalSys = new DAL(config.DbaseConStrSys);

                //check dbase connection                
                if (!dalLocal.IsConnectionOK())
                {
                    logger.Error("Connection to local database failed. " + dalLocal.ErrorMessage);
                    return false;
                }                

                //check dbase connection                
                if (!dalLocal.IsConnectionOK())
                {
                    logger.Error("Connection to sys database failed. " + dalSys.ErrorMessage);
                    return false;
                }
            }
            catch (Exception ex)
            {
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

        private static bool ProcessEODData()
        {
            //string reportDate = System.IO.File.ReadAllText(reportDateFile);
            string reportDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");

            EOD eod = new EOD(config, reportDate);
            if (eod.GenerateEndOfDay())
            {
                return true;
            }

            //Reports report = new Reports();
            //report.GenerateReportv2(config,logger);


            return true;
        }


    }
}
