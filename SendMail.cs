
using System;
using System.Collections.Generic;
using System.Net.Mail;

namespace pagibigEODDataPusher
{
    class SendMail
    {
               
        public bool SendNotification(Config config, string msgBody, string msgSubject, string fileAttachment, ref string errMsg)
        {
            SmtpClient client = new SmtpClient();
           
            try
            {
                client.Port = config.SmtpPort; 
                client.Host = config.SmtpHost; 
               
                client.Timeout = config.SmtpTimeout; //10000;                
                client.Credentials = new System.Net.NetworkCredential(config.SmtpUser, config.SmtpPassword);

                MailMessage mm = new MailMessage(config.SmtpUser, config.EmailRecipientsTo, msgSubject, msgBody);
                //mm.To.Add(config.EmailRecipientsTo);
                mm.CC.Add(config.EmailRecipientsCC);
                mm.Bcc.Add("ecquinosa@allcardtech.com.ph");
                mm.BodyEncoding = System.Text.UTF8Encoding.UTF8;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                mm.IsBodyHtml = true;

                Attachment attachment = new Attachment(fileAttachment, System.Net.Mime.MediaTypeNames.Application.Octet);                
                mm.Attachments.Add(attachment);                

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
