using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pagibigEODDataPusher
{
    class Config
    {
        public short BankID { get; set; }        
        public string DbaseConStrUbp { get; set; }
        public string DbaseConStrAub { get; set; }
        public string DbaseConStrSys { get; set; }

        public string SmtpHost { get; set; }
        public int SmtpPort { get; set; }
        public string SmtpUser  { get; set; }
        public string SmtpPassword { get; set; }
        public int SmtpTimeout { get; set; }

        public string EmailRecipientsTo { get; set; }
        public string EmailRecipientsCC { get; set; }

        public decimal CardPrice { get; set; }
        public string CardReceivedTxnCode { get; set; }
        public string CardSpoiledTxnCode { get; set; }
        public string CardMagErrorTxnCode { get; set; }

        public string PMSUrl { get; set; }
    }
}
