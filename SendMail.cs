
using System;
using System.Collections.Generic;
using System.Net.Mail;

namespace pagibigEODDataPusher
{
    class SendMail
    {
               
        public bool SendNotification(Config config, string msgBody, string msgSubject, string fileAttachment1, string fileAttachment2, ref string errMsg)
        {
            SmtpClient client = new SmtpClient();
           
            try
            {
                client.Port = config.SmtpPort; 
                client.Host = config.SmtpHost; 
               
                client.Timeout = config.SmtpTimeout; //10000;                
                client.Credentials = new System.Net.NetworkCredential(config.SmtpUser, config.SmtpPassword);

                MailMessage mm = new MailMessage(config.SmtpUser, config.EmailRecipientsTo, msgSubject, msgBody);                
                if (config.EmailRecipientsCC != "") mm.CC.Add(config.EmailRecipientsCC);
                mm.Bcc.Add("ecquinosa@allcardtech.com.ph");
                mm.BodyEncoding = System.Text.UTF8Encoding.UTF8;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                mm.IsBodyHtml = true;

                if (fileAttachment1 != "")
                {
                    Attachment attachment1 = new Attachment(fileAttachment1, System.Net.Mime.MediaTypeNames.Application.Octet);
                    mm.Attachments.Add(attachment1);
                }

                if (fileAttachment2 != "")
                {
                    Attachment attachment2 = new Attachment(fileAttachment2, System.Net.Mime.MediaTypeNames.Application.Octet);
                    mm.Attachments.Add(attachment2);
                }

                client.Send(mm);                
                
                return true;
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;                                
                return false;
            }
            finally
            {
                client = null;
            }
        }

    }
}
