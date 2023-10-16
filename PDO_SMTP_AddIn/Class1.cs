using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;

namespace PDO_SMTP_AddIn
{
    public class SendMail_using_SMTP
    {
        // PDO: Als COM-Addin erstellt. Kann dann als AddIn in VBA eingebunden und genutzt werden.

        // Hintergrund: Von VBA aus wird Stand 2023 meist die cdo-Bibliothek genutzt, welche allerdings
        // von neueren Exchange-Servern nicht mehr unterstützt wird.

        // GUID:  0b37731e-4fe7-4b76-826b-ce34e7d5b5aa

        private string Version = "20231016";
        private string strResult;



        // E-Mail-Konfiguration
        private MailMessage mail = new MailMessage();
        private string SetsmtpServer = "";
        //private int    SetsmtpPort = 587; // ohne ssl:  465; // 587;  Wird unten anhand SetIsSSL gesetzt
        private bool SetIsSSL = true;
        private string SetPassword = "";

        private string SetEmailFrom = "";  // default, sollte bei Aufruf überschrieben werden
        //private string SetemailTo = "nick@pdoetjen.de";    // Über Methode To_Add
        
        private string SetSubject = "Neue Test-E-Mail";
        private string Sethtmlbody = "Dies ist eine Test-E-Mail von mir .";



        // Werte setzen
        public void smtpServer(string strSMTPServer)
        {
            SetsmtpServer = strSMTPServer;
        }

        public void IsSSL(bool boo)
        {
            SetIsSSL = boo;
        }

        public void Password(string str)
        {
            SetPassword = str;
        }

        public void Subject(string strSubject)
        {
            SetSubject = strSubject;
        }

        public void EmailFrom(string str)
        {
            SetEmailFrom = str;
        }

        public void To_Add(string str)
        {
            mail.To.Add(str);
        }

        public void CC_Add(string str)
        {
            mail.CC.Add(str);
        }

        public void BCC_Add(string str)
        {
            mail.Bcc.Add(str);
        }

        public void HTMLBody(string str)
        {
            Sethtmlbody = str;
        }

        public void Attachment_Add(string str)
        {
            Attachment attachment = new Attachment(str);
            mail.Attachments.Add(attachment);
        }



        // Nur ein Test, ob die Verbindung zum COM steht
        public string HelloWorld()
        {
            return "Hello World, this is PDO_SMTP_AddIn Version " + Version;
        }


        public string SendMail()
        {


            // E-Mail erstellen, einige Werte hier, mit default, nicht über Methode
            // MailMessage mail = new MailMessage(emailFrom, emailTo, subject, body);
            mail.From = new MailAddress(SetEmailFrom);
            //mail.To.Add (SetemailTo);
            //mail.CC.Add(Mail_CC);
            mail.Subject = SetSubject;
            mail.Body = Sethtmlbody;
            mail.IsBodyHtml = true;

            /*Für Attachments*/
            //System.Net.Mail.Attachment attachment;
            //attachment = new System.Net.Mail.Attachment(item);
            //mail.Attachments.Add(attachment);


            // SMTP-Client erstellen und konfigurieren
            SmtpClient smtpClient = new SmtpClient(SetsmtpServer);
            smtpClient.Port =  ( SetIsSSL ) ? 587 : 465 ;//     SetsmtpPort;
            smtpClient.EnableSsl = SetIsSSL; // Aktivieren Sie SSL, wenn Ihr Server dies erfordert
            smtpClient.Credentials = new NetworkCredential(SetEmailFrom, SetPassword);
            

            // Expect100Continue auf false setzen
            ServicePointManager.Expect100Continue = true;  // wenn eigener Server, und es schneller gehen darf: false;  Das true ist wichtig bei strato

            // TLS 1.2 als Sicherheitsprotokoll verwenden
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;


            // E-Mail senden
            try
            {
                smtpClient.Send(mail);
                Console.WriteLine("E-Mail wurde erfolgreich gesendet!");
                strResult = "Gesendet";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fehler beim Senden der E-Mail: " + ex.Message);
                strResult = "Fehler beim Senden: " + ex.Message;
            }

            // Aufräumen
            mail.Dispose();
            smtpClient.Dispose();

            return strResult;

        }
    }
}


