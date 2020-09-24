using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using NLog.LogReceiverService;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace pagibigEODDataPusher
{

    class EOD
    {
        public bool isSuccess { get; set; }
        public string errorMessage { get; set; }

        private string reportDate;
        private string dateToday;

        private static DAL dalLocal = null;
        private static DAL dalSys = null;

        public EOD(string reportDate, string dateToday)
        {
            Program.logger.Info("Report date: " + reportDate);
            Program.logger.Info("Date today: " + dateToday);

            this.reportDate = reportDate;
            this.dateToday = dateToday;
            if (Program.config.BankID == Convert.ToInt16(Program.bankID.UBP)) dalLocal = new DAL(Program.config.DbaseConStrUbp);
            else dalLocal = new DAL(Program.config.DbaseConStrAub);
            dalSys = new DAL(Program.config.DbaseConStrSys);
        }

        public bool GenerateEndOfDay()
        {
            try
            { DataTable dtDailyCapturedDates = null;
                if (!dalLocal.SelectDailyCapturedByEntryDate(reportDate))
                {
                    Program.logger.Error("Failed to get DailyCapturedByEntryDate. Error " + dalLocal.ErrorMessage);
                    return false;
                }
                else dtDailyCapturedDates = dalLocal.TableResult;

                if (dtDailyCapturedDates != null)
                {
                    foreach (DataRow rw in dtDailyCapturedDates.Rows)
                    {
                        if (Convert.ToDateTime(rw["ApplicationDate"]).Date != Convert.ToDateTime(rw["EntryDate"]).Date) Program.logger.Info(string.Format("Late upload application date {0} entry date {1} total {2}", Convert.ToDateTime(rw["ApplicationDate"]).ToString("yyyy-MM-dd"), Convert.ToDateTime(rw["entryDate"]).ToString("yyyy-MM-dd"), rw["cnt"].ToString()));
                        ProcessReport(Convert.ToDateTime(rw["ApplicationDate"]).ToString("yyyy-MM-dd"), Convert.ToDateTime(rw["entryDate"]).ToString("yyyy-MM-dd"));
                    }
                }
            }
            catch (Exception ex)
            {
                Program.logger.Error("Catched runtime error " + ex.Message);
                return false;
            }

            return true;
        }

        public bool ProcessReport(string applicationDate, string entryDate)
        {
            try
            {
                DataTable dtEODdata = null;
                DataTable dtMemberRefNum = null;
                DataTable dtCardTxn = null;
                //DataTable dtCardTxnSpoiled = null;
                //DataTable dtCardTxnMagError = null;

                //if (!dalSys.CheckIfReportDateExist(reportDate))
                //{
                //    Console.Write("Failed to get CheckIfReportDateExist. Error " + dalSys.ErrorMessage);
                //    Program.logger.Error("Failed to get CheckIfReportDateExist. Error " + dalSys.ErrorMessage);
                //    return false;
                //}
                //else
                //{
                //    if (Convert.ToInt32(dalSys.ObjectResult) > 0)
                //    {
                //        Console.Write("End of day report for " + reportDate + " has been generated already");
                //        Program.logger.Error("End of day report for " + reportDate + " has been generated already");
                //        return false;
                //    }
                //}

                bool IsGetSpoiledAndMagCnt = true;               

                //do not get spoiled and mag if appdate is equal to today
                //if (Convert.ToDateTime(applicationDate).Date == Convert.ToDateTime(dateToday).Date) IsGetSpoiledAndMagCnt = false;
                if (Convert.ToDateTime(applicationDate).Date != Convert.ToDateTime(entryDate).Date) IsGetSpoiledAndMagCnt = false; //do not get spoiled and mag if appdate is not equal to entryDate

                if (!dalSys.SelectEOD_MemberRefNum_Sys(applicationDate, Program.config.BankID.ToString()))
                {
                    Program.logger.Error("Failed to get EOD_MemberRefNum_Sys. Error " + dalSys.ErrorMessage);
                    return false;
                }

                else dtMemberRefNum = dalSys.TableResult;

                int totalIssued = 0;
                //int totalReceived = 0;
                //int totalMagError = 0;
                //int totalSpoiled = 0;

                if (IsGetSpoiledAndMagCnt)
                {
                    if (!dalLocal.SelectDCS_Card_Transaction_ByEntryDate(reportDate))
                    {
                        Program.logger.Error("Failed to get DCS_Card_Transaction_ByEntryDate. Error " + dalLocal.ErrorMessage);
                        return false;
                    }
                    else dtCardTxn = dalLocal.TableResult;

                    //if (!dalLocal.SelectDCS_Card_Transaction_Spoiled_Bank(reportDate))
                    //{
                    //    Program.logger.Error("Failed to get DCS_Card_Transaction_Bank. Error " + dalLocal.ErrorMessage);
                    //    return false;
                    //}
                    //else dtCardTxnSpoiled = dalLocal.TableResult;

                    //if (!dalLocal.SelectDCS_Card_Transaction_MagError_Bank(reportDate))
                    //{
                    //    Program.logger.Error("Failed to get DCS_Card_Transaction_Bank. Error " + dalLocal.ErrorMessage);
                    //    return false;
                    //}
                    //else dtCardTxnMagError = dalLocal.TableResult;
                }

                if (!dalLocal.SelectEODData_Bank(Program.config.BankID.ToString(), reportDate))
                {

                    Program.logger.Error("Failed to get EODData_Bank. Error " + dalLocal.ErrorMessage);
                    return false;
                }
                else dtEODdata = dalLocal.TableResult;

                Program.logger.Info(string.Format("Application date {0} eod data count {1}", applicationDate, dtEODdata.DefaultView.Count.ToString()));

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

                //int? balanceCard = null;

                //foreach (var item in grpBranchWorkplaceId)
                //{
                //    totalIssued += item.totalCount;
                //}

                foreach (var item in grpBranchWorkplaceId)
                {
                    decimal expectedCash = Convert.ToDecimal((item.totalCount - item.ww) * Program.config.CardPrice);

                    int receiveCard = 0;
                    int spoiledCard = 0;
                    int magError = 0;

                    if (IsGetSpoiledAndMagCnt)
                    {
                        //check cardtxn table of equivalent branch
                        //if (dtCardTxnSpoiled.Select(string.Format("BranchCode='{0}'", item.reqBranch.ToLower())).Length > 0) spoiledCard = Convert.ToInt32(dtCardTxnSpoiled.Select(string.Format("BranchCode='{0}'", item.reqBranch.ToLower()))[0][1]);
                        //if (dtCardTxnMagError.Select(string.Format("BranchCode='{0}'", item.reqBranch.ToLower())).Length > 0) magError = Convert.ToInt32(dtCardTxnMagError.Select(string.Format("BranchCode='{0}'", item.reqBranch.ToLower()))[0][1]);

                        //if (dtCardTxn.Select(string.Format("BranchCode='{0}' AND TransactionTypeID IN ('{1}') AND WorkPlace={2}", item.reqBranch.ToLower(), Program.config.CardReceivedTxnCode.Replace(",", "','"), item.WorkplaceId.ToString())).Length > 0) receiveCard = Convert.ToInt32(dtCardTxn.Select(string.Format("BranchCode='{0}' AND TransactionTypeID IN ('{1}') AND WorkPlace={2}", item.reqBranch.ToLower(), Program.config.CardReceivedTxnCode.Replace(",", "','"), item.WorkplaceId.ToString()))[0]["Cnt"]);
                        //if (dtCardTxn.Select(string.Format("BranchCode='{0}' AND TransactionTypeID IN ('{1}') AND WorkPlace={2}", item.reqBranch.ToLower(), Program.config.CardSpoiledTxnCode.Replace(",", "','"), item.WorkplaceId.ToString())).Length > 0) spoiledCard = Convert.ToInt32(dtCardTxn.Select(string.Format("BranchCode='{0}' AND TransactionTypeID IN ('{1}') AND WorkPlace={2}", item.reqBranch.ToLower(), Program.config.CardSpoiledTxnCode.Replace(",", "','"), item.WorkplaceId.ToString()))[0]["Cnt"]);
                        //if (dtCardTxn.Select(string.Format("BranchCode='{0}' AND TransactionTypeID IN ('{1}') AND WorkPlace={2}", item.reqBranch.ToLower(), Program.config.CardMagErrorTxnCode.Replace(",", "','"), item.WorkplaceId.ToString())).Length > 0) magError = Convert.ToInt32(dtCardTxn.Select(string.Format("BranchCode='{0}' AND TransactionTypeID IN ('{1}') AND WorkPlace={2}", item.reqBranch.ToLower(), Program.config.CardMagErrorTxnCode.Replace(",", "','"), item.WorkplaceId.ToString()))[0]["Cnt"]);

                        receiveCard = GetTableCardTxnValueByBranchTxnTypeAndWorkplace(dtCardTxn, item.reqBranch.ToLower(), Program.config.CardReceivedTxnCode.Replace(",", "','"), item.WorkplaceId.ToString());
                        spoiledCard = GetTableCardTxnValueByBranchTxnTypeAndWorkplace(dtCardTxn, item.reqBranch.ToLower(), Program.config.CardSpoiledTxnCode.Replace(",", "','"), item.WorkplaceId.ToString());
                        magError = GetTableCardTxnValueByBranchTxnTypeAndWorkplace(dtCardTxn, item.reqBranch.ToLower(), Program.config.CardMagErrorTxnCode.Replace(",", "','"), item.WorkplaceId.ToString());                        

                        //////only workplaceid=1 should have spoiledcard value
                        //if (item.WorkplaceId == (short)Program.workplaceId.Deployment)
                        //{
                        //    if (grpBranchWorkplaceId.Where(t => t.reqBranch.ToString() == item.reqBranch.ToString() && Convert.ToInt32(t.WorkplaceId) == (short)Program.workplaceId.Onsite).Count() > 0)
                        //    {
                        //        //receiveCard = 0;
                        //        //spoiledCard = 0;
                        //        //magError = 0;
                        //    }                                
                        //}
                    }                    

                    if (!dalSys.Add_EodDeposits(reportDate, item.reqBranch.Trim(), item.Branch.Trim(), Program.config.BankID.ToString(), item.WorkplaceId.ToString(), receiveCard, item.totalCount, item.ww, item.nw, spoiledCard, magError, 0, expectedCash, 0, expectedCash, 0))
                    {
                        Program.logger.Error(string.Format("reqBranch {0} Branch {1} WorkplaceId {2}. Failed to add EodDeposits. Error {3}", item.reqBranch, item.Branch, item.WorkplaceId.ToString(), dalSys.ErrorMessage));
                    }                  
                }


                if (!dalSys.SelectEODDepositsByDateAndBank(Program.config.BankID.ToString(), reportDate))
                {
                    Program.logger.Error("Failed to get EODDepositsByDateAndBank. Error " + dalSys.ErrorMessage);
                    return false;
                }
                else
                {
                    foreach (DataRow rw in dalSys.TableResult.Rows)
                    {
                        //get previous data
                        //int prevBalanceCard = 0;
                        int endBalanceCard = 0;
                        //if (dalSys.SelectEODDepositsPreviousBalanceCard(Program.config.BankID.ToString(), rw["requesting_branchcode"].ToString().Trim(), Convert.ToDateTime(reportDate).AddDays(-1).ToString("yyyy-MM-dd")))
                        //{
                        //    //check first if dtResult have result
                        //    if (dalSys.TableResult.DefaultView.Count > 0)
                        //    {
                        //        if (dalSys.TableResult.Select("WorkplaceID=" + (short)Program.workplaceId.Onsite).Length > 0) prevBalanceCard = Convert.ToInt32(dalSys.TableResult.Select("WorkplaceID=" + (short)Program.workplaceId.Onsite)[0][1]);
                        //        else if (dalSys.TableResult.Select("WorkplaceID=" + (short)Program.workplaceId.Deployment).Length > 0) prevBalanceCard = Convert.ToInt32(dalSys.TableResult.Select("WorkplaceID=" + (short)Program.workplaceId.Deployment)[0][1]);
                        //    }
                        //}

                        if (dalSys.Get_ConsumablesBalance(Program.config.BankID.ToString(), rw["requesting_branchcode"].ToString().Trim(), reportDate, Program.consumableId.CardNew))
                        {
                            endBalanceCard = (int)dalSys.TableResult.Rows[0]["EndBalance"];
                        }

                        //int curBalanceCard = prevBalanceCard - ((int)rw["Issued_Card"] + (int)rw["Spoiled_Card"] + (int)rw["MagError_Card"]) + (int)rw["Received_Card"];
                        if (!dalSys.UpdateEODDepositsBalanceCard(reportDate, Program.config.BankID.ToString(), rw["requesting_branchcode"].ToString().Trim(), endBalanceCard))
                        {
                            Program.logger.Error("Failed to UpdateEODDepositsBalanceCard. Error " + dalSys.ErrorMessage);
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.logger.Error("Catched runtime error " + ex.Message);
                return false;
            }

            return true;
        }

        private static int GetTableCardTxnValueByBranchTxnTypeAndWorkplace(DataTable dtCardTxn, string reqBranch, string txnTypeCodes, string workPlaceId)
        {
            int intValue = 0;
            if (dtCardTxn == null) return intValue;

            if (dtCardTxn.Select(string.Format("BranchCode='{0}' AND TransactionTypeID IN ('{1}') AND WorkPlace={2}", reqBranch, txnTypeCodes, workPlaceId)).Length > 0)
            {
                foreach (DataRow rw in dtCardTxn.Select(string.Format("BranchCode='{0}' AND TransactionTypeID IN ('{1}') AND WorkPlace={2}", reqBranch, txnTypeCodes, workPlaceId))) intValue += Convert.ToInt32(rw["Cnt"]);
            }

            return intValue;
        }

        public static int GetGracePeriod(Program.workplaceId workplaceId, DateTime txnDate)
        {
            int noOfDays = 0;
            switch (txnDate.DayOfWeek)
            {
                case DayOfWeek.Thursday:
                    if (workplaceId == Program.workplaceId.Onsite) noOfDays = 1; else noOfDays = 4;                 
                    break;
                case DayOfWeek.Friday:
                    if (workplaceId == Program.workplaceId.Onsite) noOfDays = 3; else noOfDays = 4;
                    break;
                case DayOfWeek.Saturday:
                    if (workplaceId == Program.workplaceId.Onsite) noOfDays = 3; else noOfDays = 4;
                    break;
                case DayOfWeek.Sunday:
                    if (workplaceId == Program.workplaceId.Onsite) noOfDays = 2; else noOfDays = 3;
                    break;
                default:
                    if (workplaceId == Program.workplaceId.Onsite) noOfDays = 1; else noOfDays = 2;
                    break;
            }            

            return noOfDays;
        }

    }
}
