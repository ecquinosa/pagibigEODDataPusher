using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.ServiceModel.Syndication;
using System.Data;

namespace pagibigEODDataPusher
{
    class Program
    {

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
            //logger.Info("Application started");
            //Console.Write(DateTime.Now.DayOfWeek);
            //Console.ReadLine();

            //return;

            //validatations
            Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Initializing...");
            if (!Init()) return;            

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

                dalLocal = new DAL(config.DbaseConStr);
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
            string reportDate = "2020-06-15";


            EOD eod = new EOD(config,reportDate);
            if (eod.GenerateEndOfDay())
            {
                return true;
            }         


            return true;
        }


    }
}
