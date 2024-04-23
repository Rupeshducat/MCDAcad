using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;

namespace MCD
{
    class MailToMCD
    {
        public static void sendMail(string Msg)
        {
            MailMessage mail = new MailMessage();
            mail.To.Add("no-buildingplan@mcd.org.in");
            mail.From = new MailAddress("Acadbuildingplan@mcd.org.in");
            mail.Subject = "Error Occured";
            mail.Body = Msg;
            mail.IsBodyHtml = true;
            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(
                mail.Body, null, "text/html");
            mail.AlternateViews.Add(htmlView);
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "172.16.192.231";//server
            smtp.Credentials = new NetworkCredential(
                "Acadbuildingplan@mcd.org.in", "Acadbuildingplan123$");
            //smtp.EnableSsl = true;
            //smtp.Port = 443;
            smtp.Credentials = (ICredentialsByHost)CredentialCache.DefaultNetworkCredentials;
            // Send the email
            smtp.Send(mail);
        }
    }
}
