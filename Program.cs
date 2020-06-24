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

        private static bool ProcessEODData()
        {
            string reportDate = "2020-06-15";
            DataTable dtEODdata = null;
            DataTable dtEODdataGroup = null;
            DataTable dtMemberRefNum = null;

            if (dalSys.SelectEOD_MemberRefNum_Sys(reportDate))
            {
                dtMemberRefNum = dalSys.TableResult;
            }

            if (dalLocal.SelectEODData_Bank(config.BankID.ToString(), reportDate))
            {
                dtEODdata = dalLocal.TableResult;
                dtEODdataGroup = dtEODdata.Clone();
            }
            else
            {
                Console.WriteLine(dalLocal.ErrorMessage);
            }

            foreach (DataRow rw in dtEODdata.Rows)
            {
                if (dtMemberRefNum.Select("refNum='" + rw["RefNum"].ToString() + "'").Length > 0) rw["WorkplaceID"] = dtMemberRefNum.Select("refNum='" + rw["RefNum"].ToString() + "'")[0]["workplaceID"];
            }

            DataTable dtBranchWorkplaceId = dtEODdata.DefaultView.ToTable(true, "reqBranch", "WorkplaceID");

            Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Grouping by branch and workplace...");
            foreach (DataRow rw in dtBranchWorkplaceId.Rows)
            {
                DataTable dtTemp = dtEODdata.Select(string.Format("reqBranch='{0}' and WorkplaceID={1}", rw[0].ToString(), rw[1].ToString())).CopyToDataTable();

                DataRow newRow = dtEODdataGroup.NewRow();
                foreach (DataColumn col in dtTemp.Columns)
                {
                    switch (col.ColumnName)
                    {
                        case "RefNum":
                        case "Application_Remarks":
                            break;
                        case "nw":
                            newRow[col.ColumnName] = dtTemp.Select(string.Format("Application_Remarks LIKE '%Non-Warranty%'")).Length;
                            break;
                        case "ww":
                            newRow[col.ColumnName] = dtTemp.Select(string.Format("Application_Remarks LIKE '%With Warranty%'")).Length;
                            break;
                        case "Expected":
                            newRow[col.ColumnName] = Convert.ToDecimal((Convert.ToInt64(newRow["totalCnt"]) - Convert.ToInt64(newRow["ww"])) * 125);
                            break;
                        case "ByDSA":
                            newRow[col.ColumnName] = newRow["Expected"];
                            break;
                        case "totalCnt":
                            newRow[col.ColumnName] = dtTemp.DefaultView.Count.ToString();
                            break;
                        default:
                            newRow[col.ColumnName] = dtTemp.Rows[0][col.ColumnName];
                            break;
                    }
                }

                dtEODdataGroup.Rows.Add(newRow);
            }

            Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Done!");

            foreach (DataRow rw in dtEODdataGroup.Rows)
            {
                if (!dalSys.Add_EodDeposits(reportDate, rw["reqBranch"].ToString(), rw["Branch"].ToString(), config.BankID.ToString(), rw["WorkplaceID"].ToString(), "0", "0", rw["ww"].ToString(), rw["nw"].ToString(), "0", "0", "0", rw["Expected"].ToString(), "0", rw["ByDSA"].ToString(), "0"))
                {
                    logger.Error(dalSys.ErrorMessage);
                }
            }


            return true;
        }


    }
}
