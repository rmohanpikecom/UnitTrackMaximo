using System;
using System.Net.Mail;
using System.Configuration;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using UnitTrackMaximo;
using System.Text.RegularExpressions;

/// <summary>
/// Summary description for Mailer
/// </summary>
public class Mailer
{
    #region isMailIdCorrect
    private static bool isMailIdCorrect(string strMailId)
    {
        //Checking the Mail Id.			
        System.Text.RegularExpressions.Regex RExp = new System.Text.RegularExpressions.Regex("\\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*");
        //.Pattern = "^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w][^_]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$"

        bool blnMailIdCorrect = RExp.IsMatch(strMailId);
        return blnMailIdCorrect;
    }
    #endregion   

    #region Attachment_Mail
    public static int Attachment_Mail(string from, string To, string subject, string MailContent, string strAttachment )
    {
        int returnValue = 0;
        try
        {
            System.Net.Mail.MailMessage mailMsg = new System.Net.Mail.MailMessage();
            mailMsg.From = new System.Net.Mail.MailAddress(from, "E-Delivery");
            //string Path = ConfigurationSettings.AppSettings.Get("FileAttachmentPath_Temp");

            To = To.Replace(";", ",");
            To = To.Replace(" ", ",");
            To = Regex.Replace(To, @"\s+", "");
            if (isMailIdCorrect(To))
            {
                mailMsg.To.Add(To);
            }
            mailMsg.Bcc.Add(ConfigurationManager.AppSettings["BccEmail"].ToString());
            mailMsg.Subject = subject;

            mailMsg.Body = MailContent;
            mailMsg.BodyEncoding = System.Text.UnicodeEncoding.GetEncoding("UTF-8");

            System.Net.Mail.AlternateView plainView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(System.Text.RegularExpressions.Regex.Replace(mailMsg.Body, @"<(.|\n)*?>", string.Empty), null, "text/plain");
            System.Net.Mail.AlternateView htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(mailMsg.Body, null, "text/html");
            mailMsg.AlternateViews.Add(plainView);
            mailMsg.AlternateViews.Add(htmlView);

            mailMsg.IsBodyHtml = true;
            
            string filePath = strAttachment;
            mailMsg.Attachments.Add(new Attachment(filePath));
            string strMailServer = ConfigurationManager.AppSettings.Get("SMTPHost");
            System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient();
            if (strMailServer != "")
            {
                smtpClient.Host = strMailServer;
            }

            try
            {
                smtpClient.Send(mailMsg);
                returnValue = 1;
            }
            catch (Exception ex)
            {
                throw ex;

            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        return returnValue;
    }
    #endregion
    
}