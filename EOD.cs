using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using NLog.LogReceiverService;
using System.Runtime.CompilerServices;

namespace pagibigEODDataPusher
{

    class EOD
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public bool isSuccess { get; set; }
        public string errorMessage { get; set; }

        private Config config;
        private string reportDate;
        private DataTable dtEODdata = null;
        private DataTable dtMemberRefNum = null;
        private DataTable dtCardTxn = null;

        private static DAL dalLocal = null;
        private static DAL dalSys = null;        

        public EOD(Config config, string reportDate)
        {
            this.reportDate = reportDate;
            this.config = config;
            dalLocal = new DAL(config.DbaseConStr);
            dalSys = new DAL(config.DbaseConStrSys);
        }

        public bool GenerateEndOfDay()
        {
            try
            {
                if (!dalSys.SelectEOD_MemberRefNum_Sys(reportDate, config.BankID.ToString()))
                {
                    logger.Error("Failed to get EOD_MemberRefNum_Sys. Error " + dalSys.ErrorMessage);
                    return false;
                }
                else dtMemberRefNum = dalSys.TableResult;

                if (!dalLocal.SelectDCS_Card_Transaction_Bank(reportDate))
                {
                    logger.Error("Failed to get DCS_Card_Transaction_Bank. Error " + dalLocal.ErrorMessage);
                    return false;
                }
                else dtCardTxn = dalLocal.TableResult;


                if (!dalLocal.SelectEODData_Bank(config.BankID.ToString(), reportDate))
                {
                    logger.Error("Failed to get EODData_Bank. Error " + dalLocal.ErrorMessage);
                    return false;
                }
                else dtEODdata = dalLocal.TableResult;

                var results = from table1 in dtEODdata.AsEnumerable()
                              join table2 in dtMemberRefNum.AsEnumerable() on table1["RefNum"] equals table2["refNum"]
                              select new
                              {
                                  reqBranch = table1["reqBranch"],
                                  Branch = table1["Branch"],
                                  WorkplaceId = table2["workplaceID"],
                                  Application_Remarks = table1["Application_Remarks"]
                              };

                var grpBranchWorkplaceId = from c in results
                                           group c by new
                                           {
                                               c.reqBranch,
                                               c.Branch,
                                               c.WorkplaceId
                                           } into grpData
                                           select new vmEOD
                                           {
                                               reqBranch = grpData.Key.reqBranch.ToString(),
                                               Branch = grpData.Key.Branch.ToString(),
                                               WorkplaceId = Convert.ToInt32(grpData.Key.WorkplaceId),
                                               totalCount = grpData.Count(),
                                               nw = results.Where(t => t.reqBranch.ToString() == grpData.Key.reqBranch.ToString()
                                                                       && Convert.ToInt32(t.WorkplaceId) == Convert.ToInt32(grpData.Key.WorkplaceId)
                                                                       && t.Application_Remarks.ToString().Contains("Non-Warranty")).Count(),
                                               ww = results.Where(t => t.reqBranch.ToString() == grpData.Key.reqBranch.ToString()
                                                                       && Convert.ToInt32(t.WorkplaceId) == Convert.ToInt32(grpData.Key.WorkplaceId)
                                                                       && t.Application_Remarks.ToString().Contains("With Warranty")).Count()
                                           };

                foreach (var item in grpBranchWorkplaceId)
                {
                    decimal expectedCash = Convert.ToDecimal((item.totalCount - item.ww) * 125);
                    int spoiledCard = 0;
                    //check cardtxn table of equivalent branch
                    if (dtCardTxn.Select(string.Format("BranchCode='{0}'", item.reqBranch.ToLower())).Length > 0) spoiledCard = Convert.ToInt32(dtCardTxn.Select(string.Format("BranchCode='{0}'", item.reqBranch.ToLower()))[0][1]);
                    //only workplaceid=1 should have spoiledcard value
                    if(item.WorkplaceId==2)
                    {
                        if (grpBranchWorkplaceId.Where(t => t.reqBranch.ToString() == item.reqBranch.ToString() && Convert.ToInt32(t.WorkplaceId) == 1).Count() > 0) spoiledCard = 0;
                    }
                    
                    if (!dalSys.Add_EodDeposits(reportDate, item.reqBranch, item.Branch, config.BankID.ToString(), item.WorkplaceId, 0, item.totalCount, item.ww, item.nw, spoiledCard, 0, 0, expectedCash, 0, expectedCash, 0))
                    {
                        logger.Error(string.Format("reqBranch {0} Branch {1} WorkplaceId {2} Error {3}", item.reqBranch, item.Branch, item.WorkplaceId.ToString(), dalSys.ErrorMessage));
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("Catched runtime error " + ex.Message);
                return false;
            }            

            return true;
        }

    }
}
