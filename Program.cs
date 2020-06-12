﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace pagibigEODDataPusher
{
    class Program
    {

        private static string configFile = AppDomain.CurrentDomain.BaseDirectory + "config";
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static Config config;

        public enum bankID
        {
            UBP = 1,
            AUB
        }

        static void Main()
        {
            logger.Info("Application started");

            //validatations
            Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Initializing...");
            if (!Init()) return;

            ProcessEODData();

            logger.Info("Application closed");
        }

        private static bool Init()
        {
            DAL dal = null;
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

                //check dbase connection
                dal = new DAL(config);
                if (!dal.IsConnectionOK())
                {
                    logger.Error("Connection to database failed. " + dal.ErrorMessage);
                    return false;
                }
                dal.Dispose();
                dal = null;

            }
            catch (Exception ex)
            {
                logger.Error("Runtime catched error " + ex.Message);
                return false;
            }

            return true;
        }

        private static bool ProcessEODData()
        {
            DAL dalLocal = new DAL(config.DbaseConStr);
            DAL dalSys = new DAL(config.DbaseConStrSys);

            return true;
        }


    }
}