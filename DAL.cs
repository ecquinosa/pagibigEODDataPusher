
using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
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

                //original script
                //sb.Append(string.Format("SELECT '{0}' AS entryDate, dbo.tbl_Member.requesting_branchcode AS reqBranch, tbl_branch.Branch, {1} As BankID, 0 As WorkplaceID, NULL, 0, COUNT_BIG(*) AS totalCnt, ", reportDate, bankID));
                //sb.Append("COUNT_BIG(CASE WHEN Application_Remarks LIKE '%With Warranty%' THEN 1 END) AS ww, COUNT_BIG(CASE WHEN Application_Remarks LIKE '%Non-Warranty%' THEN 1 END) AS nw, 0 As Spoiled, 0 As MagError, ");
                //sb.Append("0 As BalanceCard, 0 As Expected, 0 As Deposited, 0 As ByDSA, 0 As ByBank, 0 As Variance, NULL As DepositoryBankID, 1, 0, 0, GETDATE(), GETDATE(), RefNum ");
                //sb.Append("FROM dbo.tbl_Member INNER JOIN ");
                //sb.Append("tbl_branch on tbl_branch.requesting_branchcode = dbo.tbl_Member.requesting_branchcode ");
                //sb.Append(string.Format("WHERE dbo.tbl_Member.EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:23:59'", reportDate));
                //sb.Append("GROUP BY dbo.tbl_Member.requesting_branchcode, tbl_branch.Branch ");
                //sb.Append("ORDER BY CAST(dbo.tbl_Member.EntryDate AS date");


                ////2nd revision
                //sb.Append(string.Format("SELECT dbo.tbl_Member.requesting_branchcode AS reqBranch, tbl_branch.Branch, {1} As BankID, 0 As WorkplaceID, 0 AS totalCnt, ", reportDate, bankID));
                //sb.Append("0 AS ww, 0 AS nw, 0 As Expected, 0 As ByDSA, 0 As ByBank, RefNum, Application_Remarks ");
                //sb.Append("FROM dbo.tbl_Member INNER JOIN ");
                //sb.Append("tbl_branch on tbl_branch.requesting_branchcode = dbo.tbl_Member.requesting_branchcode ");
                //sb.Append(string.Format("WHERE dbo.tbl_Member.EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:23:59'", reportDate));

                sb.Append(string.Format("SELECT dbo.tbl_Member.requesting_branchcode AS reqBranch, tbl_branch.Branch, RefNum, Application_Remarks ", reportDate, bankID));
                sb.Append("FROM dbo.tbl_Member INNER JOIN ");
                sb.Append("tbl_branch on tbl_branch.requesting_branchcode = dbo.tbl_Member.requesting_branchcode ");
                sb.Append(string.Format("WHERE dbo.tbl_Member.EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:23:59'", reportDate));

                //sb.Append("GROUP BY dbo.tbl_Member.requesting_branchcode, tbl_branch.Branch ");
                //sb.Append("ORDER BY CAST(dbo.tbl_Member.EntryDate AS date");

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


        public bool SelectDCS_Card_Transaction_Spoiled_Bank(string reportDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("select BranchCode, SUM(Quantity) As Spoiled from tbl_DCS_Card_Transaction ");
                sb.Append(string.Format("where TransactionTypeID IN ('03','06','07','08') and EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:23:59' ", reportDate));
                sb.Append("GROUP BY BranchCode");

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

        public bool SelectDCS_Card_Transaction_MagError_Bank(string reportDate)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("select BranchCode, SUM(Quantity) As Spoiled from tbl_DCS_Card_Transaction ");
                sb.Append(string.Format("where TransactionTypeID IN ('03','06','07','08') and EntryDate BETWEEN '{0} 00:00:00' AND '{0} 23:23:59' ", reportDate));
                sb.Append("GROUP BY BranchCode");

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

        public bool SelectDailyMonitoringReport()
        {
            try
            {
                string reportDate = DateTime.Now.ToString("yyyy-MM-dd");

                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT MONTH(Report_Date) As ReportMonth, WorkplaceID as WorkplaceId, SUM(Received_Card) as Received_Card, SUM(Issued_Card) as Issued_Card, ");
                sb.Append("SUM(WWarranty_Card) as WWarranty_Card, SUM(NWarranty_Card) as NWarranty_Card, SUM(Spoiled_Card) as Spoiled_Card, SUM(MagError_Card) as MagError_Card,  ");
                sb.Append("SUM(Balance_Card) as Balance_Card, SUM(Expected_Cash) as Expected_Cash, SUM(Deposited_Cash) as Deposited_Cash,  ");
                sb.Append("SUM(ByDSA_Cash) as ByDSA_Cash, SUM(ByBank_Cash) as ByBank_Cash, SUM(Variance) as Variance, 0 as ConsumablesUsedRibbon, 0 as ConsumablesUsedOR, 0 as ConsumablesUsedCL ");
                sb.Append("FROM dbo.tbl_DCS_EODDeposits ");
                sb.Append(string.Format("WHERE Report_Date BETWEEN '2020-01-01' AND '{0}' ", reportDate));
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
                sb.Append(string.Format("WHERE posted_date BETWEEN '{0} 00:00:00' AND '{0} 23:23:59' and bankID = {1}", reportDate, bankId));

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

        public bool Add_EodDeposits(string Report_Date, string requesting_branchcode, string Branch, string BankID, int WorkplaceID, int Received_Card, int Issued_Card, int WWarranty_Card,
                                    int NWarranty_Card, int Spoiled_Card, int MagError_Card, int Balance_Card, decimal Expected_Cash, decimal Deposited_Cash, decimal ByDSA_Cash, decimal ByBank_Cash)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("INSERT INTO dbo.tbl_DCS_EODDeposits (Report_Date, requesting_branchcode, Branch, BankID, WorkplaceID, Username, Received_Card, Issued_Card, WWarranty_Card, NWarranty_Card, Spoiled_Card, MagError_Card, Balance_Card, Expected_Cash, Deposited_Cash, ");
                sb.Append("ByDSA_Cash, ByBank_Cash, Variance, DepositoryBankID, StatusTypeID, ReworkCntr, ExcessAppForm, Posted_Date, LastUpdated_Date) ");
                sb.Append(" VALUES ");
                sb.Append("(@Report_Date, @requesting_branchcode, @Branch, @BankID, @WorkplaceID, NULL, @Received_Card, @Issued_Card, @WWarranty_Card, @NWarranty_Card, @Spoiled_Card, @MagError_Card, @Balance_Card, @Expected_Cash, @Deposited_Cash, ");
                sb.Append("@ByDSA_Cash, @ByBank_Cash, 0, NULL, 1, 0, 0, GETDATE(), GETDATE()); SELECT SCOPE_IDENTITY() AS [SCOPE_IDENTITY];  ");



                OpenConnection();
                cmd = new SqlCommand(sb.ToString(), con);
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
                cmd.Parameters.AddWithValue("Balance_Card", Balance_Card);
                cmd.Parameters.AddWithValue("Expected_Cash", Expected_Cash);
                cmd.Parameters.AddWithValue("Deposited_Cash", Deposited_Cash);
                cmd.Parameters.AddWithValue("ByDSA_Cash", ByDSA_Cash);
                cmd.Parameters.AddWithValue("ByBank_Cash", ByBank_Cash);

                _ExecuteScalar(CommandType.Text);

                return true;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
                return false;
            }
        }

        public bool Add_EODDeployed(int EODDepositID, int workplaceId)
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
                if(workplaceId==1)
                    cmd.Parameters.AddWithValue("DaysGracePeriod", 1);                
                else
                    cmd.Parameters.AddWithValue("DaysGracePeriod", 2);

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
