
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.ServiceModel.Configuration;
using System.Text;

namespace pagibigEODDataPusher
{
    class DAL : IDisposable

    {

        private DataTable dtResult;
        //private DataSet dsResult;
        private object objResult;
        private IDataReader _readerResult;
        private string strErrorMessage;

        private SqlConnection con;
        private SqlCommand cmd;
        private SqlDataAdapter da;

        private string ConStr { get; set; }

        public string ErrorMessage { get { return strErrorMessage; } }

        public DataTable TableResult { get { return dtResult; } }

        public object ObjectResult { get { return objResult; } }

        public DAL(string ConStr)
        {
            this.ConStr = ConStr;
        }

        public void ClearAllPools()
        {
            SqlConnection.ClearAllPools();
        }

        private void OpenConnection()
        {
            if (con == null) con = new SqlConnection(ConStr);
        }

        private void CloseConnection()
        {
            if (cmd != null) cmd.Dispose();
            if (da != null) da.Dispose();
            if (_readerResult != null)
            {
                _readerResult.Close();
                _readerResult.Dispose();
            }
            if (con != null)
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
            ClearAllPools();
        }

        private void ExecuteNonQuery(CommandType cmdType)
        {
            cmd.CommandType = cmdType;

            // If con.State = ConnectionState.Open Then con.Close()
            // con.Open()
            if (con.State == ConnectionState.Closed)
                con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        private void _ExecuteScalar(CommandType cmdType)
        {
            cmd.CommandType = cmdType;

            // If con.State = ConnectionState.Open Then con.Close()
            // con.Open()
            if (con.State == ConnectionState.Closed) con.Open();
            object _obj;
            _obj = cmd.ExecuteScalar();
            con.Close();

            objResult = _obj;
        }

        private void _ExecuteReader(CommandType cmdType)
        {
            cmd.CommandType = cmdType;

            // If con.State = ConnectionState.Open Then con.Close()
            // con.Open()
            if (con.State == ConnectionState.Closed)
                con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            _readerResult = reader;
        }

        private void FillDataAdapter(CommandType cmdType)
        {
            cmd.CommandTimeout = 0;
            cmd.CommandType = cmdType;
            da = new SqlDataAdapter(cmd);
            DataTable _dt = new DataTable();
            da.Fill(_dt);
            dtResult = _dt;
        }

        public bool SelectQuery(string strQuery)
        {
            try
            {
                OpenConnection();
                cmd = new SqlCommand(strQuery, con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectEODData_Bank(string bankID, string reportDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Append(string.Format("SELECT dbo.tbl_Member.requesting_branchcode AS reqBranch, tbl_branch.Branch, RefNum, Application_Remarks ", reportDate, bankID));
                sb.Append("FROM dbo.tbl_Member INNER JOIN ");
                sb.Append("tbl_branch on tbl_branch.requesting_branchcode = dbo.tbl_Member.requesting_branchcode ");
                sb.Append(string.Format("WHERE dbo.tbl_Member.EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:59:59'", reportDate));                

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectDailyCapturedByEntryDate(string reportDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Append("select ApplicationDate, cast(EntryDate as date) as EntryDate, count(ApplicationDate) as cnt from tbl_Member ");
                sb.Append(" GROUP BY ApplicationDate, cast(EntryDate as date) ");
                sb.Append(string.Format(" HAVING cast(EntryDate as date) = '{0}'",reportDate));
                sb.Append(" ORDER BY ApplicationDate ");                

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectLastTwoEntryDates()
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Append("select distinct top 2 cast(entryDate as date) as last2dates from tbl_Member ");
                sb.Append("order by cast(entryDate as date) desc ");                

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectEODDepositsPreviousBalanceCard(string bankId, string reqBranch, string reportDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                
                sb.Append(string.Format("SELECT WorkplaceID, Balance_Card FROM dbo.tbl_DCS_EODDeposits"));                
                //sb.Append(string.Format(" WHERE Report_Date BETWEEN '{0}' AND '{1}' AND BankID={2} AND WorkplaceID={3} AND requesting_branchcode='{4}'", Convert.ToDateTime(reportDate).AddDays(-30).ToString("yyyy-MM-dd"), reportDate, bankId, 1, reqBranch));
                sb.Append(string.Format(" WHERE Report_Date BETWEEN '{0}' AND '{1}' AND BankID={2} AND requesting_branchcode='{3}'", Convert.ToDateTime(reportDate).AddDays(-30).ToString("yyyy-MM-dd"), reportDate, bankId, reqBranch));
                sb.Append(string.Format(" ORDER BY Report_Date DESC"));                

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectEODDepositsLastBalanceCardByDateAndBankIdAndReBranch(string bankId, string reqBranch, string reportDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Append(string.Format("SELECT TOP 1 Balance_Card FROM dbo.tbl_DCS_EODDeposits"));
                //sb.Append(string.Format(" WHERE Report_Date BETWEEN '{0}' AND '{1}' AND BankID={2} AND WorkplaceID={3} AND requesting_branchcode='{4}'", Convert.ToDateTime(reportDate).AddDays(-30).ToString("yyyy-MM-dd"), reportDate, bankId, 1, reqBranch));
                sb.Append(string.Format(" WHERE Report_Date BETWEEN '{0}' AND '{1}' AND BankID={2} AND requesting_branchcode='{3}'", Convert.ToDateTime(reportDate).AddDays(-30).ToString("yyyy-MM-dd"), reportDate, bankId, reqBranch));
                sb.Append(string.Format(" ORDER BY Report_Date DESC"));

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectEODDepositsDistinctBranchByDateAndBankId(string bankId, string startDate, string endDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();                

                sb.Append(string.Format("SELECT DISTINCT requesting_branchcode FROM dbo.tbl_DCS_EODDeposits "));
                sb.Append(string.Format(" WHERE Report_Date BETWEEN '{0}' AND '{1}' AND BankID={2}", startDate, endDate, bankId));
           

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectEODDepositsByDateAndBank(string bankId, string reportDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Append(string.Format("select requesting_branchcode, SUM(Received_Card) Received_Card, SUM(Issued_Card) Issued_Card, SUM(Spoiled_Card) Spoiled_Card, SUM(MagError_Card) MagError_Card from tbl_DCS_EODDeposits "));
                sb.Append(string.Format("where Report_Date = '{0}' AND BankID = {1} ", reportDate, bankId));
                sb.Append(string.Format("GROUP BY requesting_branchcode "));                

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectDCS_Card_Transaction_ByEntryDate(string reportDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                string txnTypes = string.Format("'{0}','{1}','{2}'",Program.config.CardReceivedTxnCode.Replace(",","','"), Program.config.CardSpoiledTxnCode.Replace(",", "','"), Program.config.CardMagErrorTxnCode.Replace(",", "','"));

                sb.Append("select BranchCode, TransactionTypeID, Workplace, SUM(Quantity) As Cnt from tbl_DCS_Card_Transaction ");
                sb.Append(string.Format("where TransactionTypeID IN ({0}) and TransactionDate BETWEEN '{1} 00:00:00' AND '{1} 23:59:59' ", txnTypes, reportDate));                
                sb.Append("GROUP BY BranchCode, TransactionTypeID, Workplace");

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectConsumables(string bankId, string startDate, string endDate, string txnTypeId)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                //commented out on 09/11
                //sb.Append("SELECT aa.ConsumableID, MONTH(dbo.tbl_DCS_EODDeposits.Report_Date) ReportMonth, SUM(cast(aa.Quantity as int)) As Total ");
                //sb.Append("FROM dbo.tbl_DCS_DSAConsumables aa INNER JOIN ");
                //sb.Append("dbo.tbl_DCS_EODDeposits ON aa.EODDepositsID = dbo.tbl_DCS_EODDeposits.ID ");
                //sb.Append(string.Format("WHERE aa.ConsumableID IN (1, 2, 3) AND dbo.tbl_DCS_EODDeposits.BankID = {0} AND aa.TransactionTypeID = {1} AND ", bankId, txnTypeId));
                //sb.Append(string.Format("dbo.tbl_DCS_EODDeposits.Report_Date BETWEEN '{0}' AND '{1}' ", startDate, endDate));
                //sb.Append("GROUP BY aa.ConsumableID, MONTH(dbo.tbl_DCS_EODDeposits.Report_Date) ");

                sb.Append("SELECT aa.ConsumableID, MONTH(aa.Posted_Date) ReportMonth, SUM(cast(aa.Quantity as int)) As Total ");
                sb.Append("FROM dbo.tbl_DCS_DSAConsumables aa ");                
                sb.Append(string.Format("WHERE aa.ConsumableID IN (1, 2, 3) AND aa.bankID = {0} AND aa.TransactionTypeID = {1} AND ", bankId, txnTypeId));
                sb.Append(string.Format("aa.Posted_Date BETWEEN '{0} 00:00:00' AND '{1} 23:59:59' ", startDate, endDate));
                sb.Append("GROUP BY aa.ConsumableID, MONTH(aa.Posted_Date) ");

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }
         
        public bool SelectConsumablesInOut(string bankId, string startDate, string endDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                //09/11
                //sb.Append("SELECT dbo.tbl_DCS_EODDeposits.requesting_branchcode, aa.TransactionTypeID, aa.ConsumableID, cast(aa.Quantity as int) Quantity ");
                //sb.Append("FROM dbo.tbl_DCS_DSAConsumables aa INNER JOIN ");
                //sb.Append("dbo.tbl_DCS_EODDeposits ON aa.EODDepositsID = dbo.tbl_DCS_EODDeposits.ID ");
                //sb.Append(string.Format("WHERE aa.ConsumableID IN (1, 2, 3) AND dbo.tbl_DCS_EODDeposits.BankID = {0} AND aa.TransactionTypeID IN (1, 2) AND ", bankId));
                //sb.Append(string.Format("dbo.tbl_DCS_EODDeposits.Report_Date BETWEEN '{0}' AND '{1}' ", startDate, endDate));

                sb.Append("SELECT aa.branchCode requesting_branchcode, aa.TransactionTypeID, aa.ConsumableID, cast(aa.Quantity as int) Quantity ");
                sb.Append("FROM dbo.tbl_DCS_DSAConsumables aa ");                
                sb.Append(string.Format("WHERE aa.ConsumableID IN (1, 2, 3) AND aa.bankID = {0} AND aa.TransactionTypeID IN (1, 2) AND ", bankId));
                sb.Append(string.Format("aa.Posted_Date BETWEEN '{0} 00:00:00' AND '{1} 23:59:59' ", startDate, endDate));

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectDailyMonitoringReport2(string bankId, string startDate, string endDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Append("SELECT dbo.tbl_DCS_EODDeposits.Branch, dbo.tbl_DCS_Bank.BankCode, dbo.tbl_Workplace.Workplace, dbo.tbl_DCS_EODDeposits.Received_Card, dbo.tbl_DCS_EODDeposits.Issued_Card, dbo.tbl_DCS_EODDeposits.WWarranty_Card,  ");
                sb.Append("dbo.tbl_DCS_EODDeposits.NWarranty_Card, dbo.tbl_DCS_EODDeposits.Spoiled_Card, dbo.tbl_DCS_EODDeposits.MagError_Card, dbo.tbl_DCS_EODDeposits.Balance_Card, dbo.tbl_DCS_EODDeposits.Expected_Cash,   ");
                sb.Append("dbo.tbl_DCS_EODDeposits.Deposited_Cash, dbo.tbl_DCS_EODDeposits.ByDSA_Cash, dbo.tbl_DCS_EODDeposits.ByBank_Cash, dbo.tbl_DCS_EODDeposits.Variance,  ");
                sb.Append("dbo.tbl_DCS_DepositBankAccount.AccountName, dbo.tbl_EODStatusType.StatusType, 0 RemRibbon, 0 RemOR, 0 RemCL, dbo.tbl_DCS_EODDeposits.requesting_branchcode  ");
                sb.Append("FROM dbo.tbl_DCS_EODDeposits LEFT OUTER JOIN  ");
                sb.Append("dbo.tbl_DCS_Bank ON dbo.tbl_DCS_EODDeposits.BankID = dbo.tbl_DCS_Bank.BankID LEFT OUTER JOIN  ");
                sb.Append("dbo.tbl_Workplace ON dbo.tbl_DCS_EODDeposits.WorkplaceId = dbo.tbl_Workplace.ID LEFT OUTER JOIN  ");
                sb.Append("dbo.tbl_DCS_EODRefDepositTxn ON dbo.tbl_DCS_EODDeposits.ID = dbo.tbl_DCS_EODRefDepositTxn.EODDepositID LEFT OUTER JOIN  ");
                sb.Append("dbo.tbl_DCS_DepositTransaction ON dbo.tbl_DCS_EODRefDepositTxn.DepositTransactionID = dbo.tbl_DCS_DepositTransaction.Id LEFT OUTER JOIN  ");
                sb.Append("dbo.tbl_DCS_DepositBankAccount ON tbl_DCS_DepositBankAccount.Id = dbo.tbl_DCS_EODDeposits.DepositoryBankID LEFT OUTER JOIN  ");
                sb.Append("dbo.tbl_EODStatusType ON dbo.tbl_EODStatusType.Id = tbl_DCS_EODDeposits.StatusTypeId  ");
                sb.Append(string.Format("WHERE dbo.tbl_DCS_EODDeposits.BankID = {0} AND dbo.tbl_DCS_EODDeposits.Report_Date BETWEEN '{1}' AND '{2}' ", bankId, startDate, endDate));
                //sb.Append(string.Format("WHERE dbo.tbl_DCS_EODDeposits.BankID = {0} AND dbo.tbl_DCS_EODDeposits.Report_Date BETWEEN '{1}' AND '{2}' ",bankId, "2020-08-18", "2020-08-18"));

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool Get_ConsumablesBalance(string bankId, string reqBranch, string reportDate, Program.consumableId consumableId)
        {
            try
            {
                OpenConnection();
                cmd = new SqlCommand("spGet_ConsumablesBalance", con);
                cmd.Parameters.AddWithValue("bankId", bankId);
                cmd.Parameters.AddWithValue("reqBranch", reqBranch);
                cmd.Parameters.AddWithValue("consumablesId", (int)consumableId);                
                cmd.Parameters.AddWithValue("endDate", reportDate);

                FillDataAdapter(CommandType.StoredProcedure);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }


        //public bool SelectDCS_Card_Transaction_Spoiled_Bank(string reportDate)
        //{
        //    try
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        sb.Append("select BranchCode, SUM(Quantity) As Spoiled from tbl_DCS_Card_Transaction ");
        //        //sb.Append(string.Format("where TransactionTypeID IN ('03','06','07','08') and EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:59:59' ", reportDate));
        //        sb.Append(string.Format("where TransactionTypeID IN ('03') and EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:59:59' ", reportDate));
        //        sb.Append("GROUP BY BranchCode");

        //        OpenConnection();
        //        cmd = new SqlCommand(sb.ToString(), con);

        //        FillDataAdapter(CommandType.Text);

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        strErrorMessage = ex.Message;
        //        return false;
        //    }
        //}

        //public bool SelectDCS_Card_Transaction_MagError_Bank(string reportDate)
        //{
        //    try
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        sb.Append("select BranchCode, SUM(Quantity) As Spoiled from tbl_DCS_Card_Transaction ");
        //        //sb.Append(string.Format("where TransactionTypeID IN ('03','06','07','08') and EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:59:59' ", reportDate));
        //        sb.Append(string.Format("where TransactionTypeID IN ('13') and EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:59:59' ", reportDate));
        //        sb.Append("GROUP BY BranchCode");

        //        OpenConnection();
        //        cmd = new SqlCommand(sb.ToString(), con);

        //        FillDataAdapter(CommandType.Text);

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        strErrorMessage = ex.Message;
        //        return false;
        //    }
        //}

        public bool SelectDailyMonitoringReport(string bankId, string startDate, string endDate)
        {
            try
            {
                string reportDate = DateTime.Now.ToString("yyyy-MM-dd");

                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT MONTH(Report_Date) As ReportMonth, WorkplaceID as WorkplaceId, SUM(Received_Card) as Received_Card, SUM(Issued_Card) as Issued_Card, ");
                sb.Append("SUM(WWarranty_Card) as WWarranty_Card, SUM(NWarranty_Card) as NWarranty_Card, SUM(Spoiled_Card) as Spoiled_Card, SUM(MagError_Card) as MagError_Card,  ");
                sb.Append("SUM(Balance_Card) as Balance_Card, SUM(Expected_Cash) as Expected_Cash, SUM(CASE WHEN StatusTypeID = 5 THEN Deposited_Cash ELSE 0 END) as Deposited_Cash,  ");
                sb.Append("SUM(CASE WHEN StatusTypeID = 5 THEN ByDSA_Cash ELSE 0 END) as ByDSA_Cash, SUM(CASE WHEN StatusTypeID = 5 THEN ByBank_Cash ELSE 0 END) as ByBank_Cash, (SUM(Expected_Cash)-SUM(CASE WHEN StatusTypeID = 5 THEN Deposited_Cash ELSE 0 END)) as Variance, 0 as ConsumablesUsedRibbon, 0 as ConsumablesUsedOR, 0 as ConsumablesUsedCL ");
                sb.Append("FROM dbo.tbl_DCS_EODDeposits ");                
                sb.Append(string.Format("WHERE Report_Date BETWEEN '{0}' AND '{1}' AND BankID = {2}", startDate, endDate, bankId));                
                sb.Append("GROUP BY MONTH(Report_Date), WorkplaceID ");                

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool SelectEOD_MemberRefNum_Sys(string reportDate, string bankId)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT refNum, workplaceID, bankID FROM tbl_EOD_MemberRefNum ");                
                sb.Append(string.Format("WHERE posted_date BETWEEN '{0} 00:00:00' AND '{1} 23:59:59' and bankID = {2}", Convert.ToDateTime(reportDate).AddDays(-5).ToString("yyyy-MM-dd"), reportDate, bankId));

                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);

                FillDataAdapter(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool IsConnectionOK(string strConString = "")
        {
            try
            {
                if (strConString != "")
                    ConStr = strConString;
                OpenConnection();

                con.Open();
                con.Close();

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool UpdateEODDepositsBalanceCard(string reportDate, string bankId, string reqBranch, int balanceCard)
        {
            try
            {
                OpenConnection();

                StringBuilder sb = new StringBuilder();
                    
                sb.Append(string.Format("update tbl_DCS_EODDeposits set Balance_Card = @Balance_Card "));
                sb.Append(string.Format("where Report_Date = '{0}' AND BankID = {1} AND requesting_branchcode = '{2}' ", reportDate, bankId, reqBranch));                

                cmd = new SqlCommand(sb.ToString(), con);

                cmd.Parameters.AddWithValue("Balance_Card", balanceCard);

                ExecuteNonQuery(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool ExecuteQuery(string strQuery)
        {
            try
            {
                OpenConnection();
                cmd = new SqlCommand(strQuery, con);

                ExecuteNonQuery(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool GenerateConsumbalesDailyEnv(string bankId, string reportDate)
        {
            try
            {
                OpenConnection();
                cmd = new SqlCommand("spGenerateConsumbalesDailyEnv", con);

                cmd.Parameters.AddWithValue("report_date", reportDate);
                cmd.Parameters.AddWithValue("bankId", bankId);                
                //cmd.Parameters.AddWithValue("startDate", balanceCard);
                cmd.Parameters.AddWithValue("endDate", reportDate);

                ExecuteNonQuery(CommandType.StoredProcedure);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool ExecuteScalar(string strQuery)
        {
            try
            {
                OpenConnection();
                cmd = new SqlCommand(strQuery, con);

                _ExecuteScalar(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }
        public bool CheckIfReportDateExist(string reportDate)
        {
            try
            {
                //string reportDate = DateTime.Now.ToString("yyyy-MM-dd");
                OpenConnection();
                cmd = new SqlCommand("select count(*) from dbo.tbl_DCS_EODDeposits" + string.Format(" where Report_Date = '{0}' ", reportDate), con);

                _ExecuteScalar(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }        

        //public bool Check_LoanDeductionIfExist(EOD eod)
        //{
        //    try
        //    {
        //        OpenConnection();
        //        cmd = new SqlCommand("SELECT COUNT(*) FROM tbl_LoanDeductionRecon WHERE PagIBIGID=@PagIBIGID AND PaymentRefNo=@PaymentRefNo AND ReferenceFile=@ReferenceFile", con);
        //        cmd.Parameters.AddWithValue("PagIBIGID", eod.PagIBIGID);
        //        cmd.Parameters.AddWithValue("PaymentRefNo", eod.PaymentRefNo);
        //        cmd.Parameters.AddWithValue("ReferenceFile", eod.ReferenceFile);

        //        _ExecuteScalar(CommandType.Text);

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        strErrorMessage = ex.Message;
        //        return false;
        //    }
        //}

        public bool Add_EodDeposits(string Report_Date, string requesting_branchcode, string Branch, string BankID, string WorkplaceID, int Received_Card, int Issued_Card, int WWarranty_Card,
                                    int NWarranty_Card, int Spoiled_Card, int MagError_Card, int Balance_Card, decimal Expected_Cash, decimal Deposited_Cash, decimal ByDSA_Cash, decimal ByBank_Cash)
                                    //int totalIssued, ref int? balanceCard)
        {
            try
            {
                StringBuilder sbInsert = new StringBuilder();
                StringBuilder sbUpdate = new StringBuilder();
                
                int id = 0;
                int? curBalanceCard = 0;

                ////get previous data
                //int prevBalanceCard = 0;
                //if (SelectEODDepositsPreviousBalanceCard(BankID, requesting_branchcode, Convert.ToDateTime(Report_Date).AddDays(-1).ToString("yyyy-MM-dd")))
                //{
                //    //check first if dtResult have result
                //    if (dtResult.DefaultView.Count > 0)
                //    {                        
                //        if (dtResult.Select("WorkplaceID=" + (short)Program.workplaceId.Onsite).Length > 0) prevBalanceCard = Convert.ToInt32(dtResult.Select("WorkplaceID=" + (short)Program.workplaceId.Onsite)[0][1]);
                //        else if (dtResult.Select("WorkplaceID=" + (short)Program.workplaceId.Deployment).Length > 0) prevBalanceCard = Convert.ToInt32(dtResult.Select("WorkplaceID=" + (short)Program.workplaceId.Deployment)[0][1]);                        
                //    }
                //}

                if (!ExecuteScalar(string.Format("SELECT ID FROM dbo.tbl_DCS_EODDeposits WHERE Report_Date='{0}' AND BankID={1} AND requesting_branchcode='{2}' AND WorkplaceID={3}", Report_Date, BankID, requesting_branchcode, WorkplaceID)))
                {
                    strErrorMessage = string.Format("Failed to select DCS_EODDeposits WHERE Report_Date='{0}' AND BankID={1} AND requesting_branchcode='{2}' AND WorkplaceID={3}", Report_Date, BankID, requesting_branchcode, WorkplaceID);
                    return false;
                }
                else if (objResult != null) id = Convert.ToInt32(objResult);                

                if (id == 0)
                {
                    sbInsert.Append("INSERT INTO dbo.tbl_DCS_EODDeposits (Report_Date, requesting_branchcode, Branch, BankID, WorkplaceID, Username, Received_Card, Issued_Card, WWarranty_Card, NWarranty_Card, Spoiled_Card, MagError_Card, Balance_Card, Expected_Cash, Deposited_Cash, ");
                    sbInsert.Append(" ByDSA_Cash, ByBank_Cash, Variance, DepositoryBankID, StatusTypeID, ReworkCntr, ExcessAppForm, Posted_Date, LastUpdated_Date) ");
                    sbInsert.Append(" VALUES ");
                    sbInsert.Append(" (@Report_Date, @requesting_branchcode, @Branch, @BankID, @WorkplaceID, NULL, @Received_Card, @Issued_Card, @WWarranty_Card, @NWarranty_Card, @Spoiled_Card, @MagError_Card, @Balance_Card, @Expected_Cash, @Deposited_Cash, ");
                    sbInsert.Append(" @ByDSA_Cash, @ByBank_Cash, @Variance, 0, 1, 0, 0, GETDATE(), GETDATE()); SELECT SCOPE_IDENTITY() AS [SCOPE_IDENTITY];  ");
                }
                else
                {
                    sbUpdate.Append("UPDATE dbo.tbl_DCS_EODDeposits");
                    //sbUpdate.Append(" SET Issued_Card=@Issued_Card, WWarranty_Card=@WWarranty_Card, NWarranty_Card=@NWarranty_Card, Balance_Card=@Balance_Card, Expected_Cash=@Expected_Cash, ");
                    sbUpdate.Append(" SET Issued_Card=@Issued_Card, WWarranty_Card=@WWarranty_Card, NWarranty_Card=@NWarranty_Card, Expected_Cash=@Expected_Cash, ");
                    sbUpdate.Append(" ByDSA_Cash=@ByDSA_Cash, ByBank_Cash=@ByBank_Cash, Variance=@Expected_Cash-Deposited_Cash, Received_Card = @Received_Card, Spoiled_Card = @Spoiled_Card, MagError_Card = @MagError_Card, LastUpdated_Date=GETDATE() ");
                    sbUpdate.Append(" WHERE ID = @ID ");
                }

                OpenConnection();
                if (id == 0) cmd = new SqlCommand(sbInsert.ToString(), con);
                else cmd = new SqlCommand(sbUpdate.ToString(), con);                

                //if (balanceCard == null)
                //{
                //    //curBalanceCard = prevBalanceCard - (Issued_Card + Spoiled_Card + MagError_Card) + Received_Card;
                //    curBalanceCard = prevBalanceCard - (totalIssued + Spoiled_Card + MagError_Card) + Received_Card;
                //    balanceCard = curBalanceCard;
                //}
                //else curBalanceCard = balanceCard;

                if (id == 0)
                {
                    cmd.Parameters.AddWithValue("Report_Date", Report_Date);
                    cmd.Parameters.AddWithValue("requesting_branchcode", requesting_branchcode);
                    cmd.Parameters.AddWithValue("Branch", Branch);
                    cmd.Parameters.AddWithValue("BankID", BankID);
                    cmd.Parameters.AddWithValue("WorkplaceID", WorkplaceID);
                    cmd.Parameters.AddWithValue("Received_Card", Received_Card);
                    cmd.Parameters.AddWithValue("Issued_Card", Issued_Card);
                    cmd.Parameters.AddWithValue("WWarranty_Card", WWarranty_Card);
                    cmd.Parameters.AddWithValue("NWarranty_Card", NWarranty_Card);
                    cmd.Parameters.AddWithValue("Spoiled_Card", Spoiled_Card);
                    cmd.Parameters.AddWithValue("MagError_Card", MagError_Card);
                    cmd.Parameters.AddWithValue("Balance_Card", curBalanceCard);
                    cmd.Parameters.AddWithValue("Expected_Cash", Expected_Cash);
                    cmd.Parameters.AddWithValue("Deposited_Cash", Deposited_Cash);
                    cmd.Parameters.AddWithValue("ByDSA_Cash", ByDSA_Cash);
                    cmd.Parameters.AddWithValue("ByBank_Cash", ByBank_Cash);
                    cmd.Parameters.AddWithValue("Variance", Expected_Cash-Deposited_Cash);

                    _ExecuteScalar(CommandType.Text);
                }
                else
                {
                    cmd.Parameters.AddWithValue("ID", id);                    
                    cmd.Parameters.AddWithValue("Issued_Card", Issued_Card);
                    cmd.Parameters.AddWithValue("WWarranty_Card", WWarranty_Card);
                    cmd.Parameters.AddWithValue("NWarranty_Card", NWarranty_Card);                    
                    //cmd.Parameters.AddWithValue("Balance_Card", curBalanceCard);
                    cmd.Parameters.AddWithValue("Expected_Cash", Expected_Cash);                    
                    cmd.Parameters.AddWithValue("ByDSA_Cash", ByDSA_Cash);
                    cmd.Parameters.AddWithValue("ByBank_Cash", ByBank_Cash);
                    cmd.Parameters.AddWithValue("Received_Card", Received_Card);
                    cmd.Parameters.AddWithValue("Spoiled_Card", Spoiled_Card);
                    cmd.Parameters.AddWithValue("MagError_Card", MagError_Card);
                    //cmd.Parameters.AddWithValue("Variance", Expected_Cash - Deposited_Cash);

                    ExecuteNonQuery(CommandType.Text);
                }                

                if (id == 0)
                {
                    if (!Add_EODDeployed(Convert.ToInt32(objResult), WorkplaceID, Report_Date))
                    {
                        Program.logger.Error(string.Format("reqBranch {0} Branch {1} WorkplaceId {2}. Failed to add EODDeployed. Error {3}", requesting_branchcode, Branch, WorkplaceID, ErrorMessage));
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool Add_EODDeployed(int EODDepositID, string workplaceId, string Report_Date)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("INSERT INTO dbo.tbl_DCS_EODDeployed (EODDepositID, DaysGracePeriod, Posted_Date, LastUpdated_Date) ");                
                sb.Append(" VALUES ");                
                sb.Append("(@EODDepositID, @DaysGracePeriod, GETDATE(), GETDATE()) ");


                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);
                cmd.Parameters.AddWithValue("EODDepositID", EODDepositID);
                //if(workplaceId=="1")
                //    cmd.Parameters.AddWithValue("DaysGracePeriod", 1);                
                //else
                //    cmd.Parameters.AddWithValue("DaysGracePeriod", 2);

                cmd.Parameters.AddWithValue("DaysGracePeriod", EOD.GetGracePeriod((Program.workplaceId)Convert.ToInt32(workplaceId), Convert.ToDateTime(Report_Date)));                

                ExecuteNonQuery(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        private bool disposed = false;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    // Dispose managed resources.
                    CloseConnection();
                }



                // Note disposing has been done.
                disposed = true;

            }
        }

    }
}
