using System;
using System.Net;
using System.Net.Mail;

namespace PDO_SMTP_AddIn
{
    public class SendMail_using_SMTP
    {
        // PDO: Configured as COM-Addin. Can be used as AddIn in VBA , after installation and inclusion as reference.

        // Background: From VBA status quo 2023 mostly the cdo-library is used, which is 
        // deprecated / not supported by newer Exchange-Servers (>= 2019).

        // GUID:  0b37731e-4fe7-4b76-826b-ce34e7d5b5aa

        private string Version = "20231016";
        private string strResult = "";



        // E-Mail-Configuration
        private MailMessage mail = new MailMessage();   // new MailMessage(emailFrom, emailTo, subject, body);
        private string SetsmtpServer = "";
        //private int    SetsmtpPort = 587; // ohne ssl:  465; // 587;  Will be set using SetIsSSL 
        private bool SetIsSSL = true;
        private string SetPassword = "";

        private string SetEmailFrom = "";  // default, sollte bei Aufruf überschrieben werden
        
        private string SetSubject = "New Test-E-Mail";
        private string Sethtmlbody = "Can I <string>test</strong> you?";



        // Methods to set values
        public void smtpServer(string str)
        {
            SetsmtpServer = str;
        }

        public void IsSSL(bool boo)
        {
            SetIsSSL = boo;
        }

        public void Password(string str)
        {
            SetPassword = str;
        }

        public void Subject(string str)
        {
            SetSubject = str;
        }

        public void EmailFrom(string str)
        {
            SetEmailFrom = str;
        }

        public void To_Add(string str)
        {
            if (IsValidEmail(str))
            {
                mail.To.Add(str);
            } else
            {
                strResult = "PDO_SMTP_AddIn: The recipient-address is invalid. Please check.";
            }
        }

        public void CC_Add(string str)
        {
            if (IsValidEmail(str))
            {
                mail.CC.Add(str);
            }
            else
            {
                strResult = "PDO_SMTP_AddIn: The CC-address is invalid. Please check.";
            }
        }

        public void BCC_Add(string str)
        {
            if (IsValidEmail(str))
            {
                mail.Bcc.Add(str);
            }
            else
            {
                strResult = "PDO_SMTP_AddIn: The Bcc-address is invalid. Please check.";
            }
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



        // Just to check for connection to COM 
        public string HelloWorld()
        {
            return "Hello World, this is PDO_SMTP_AddIn Version " + Version;
        }


        // Main function
        public string SendMail()
        {


            // Configure Mail with some defaults, not using the methods, to make it more tolerant to faults
            
            if (string.IsNullOrEmpty(SetEmailFrom))
            {
                throw new Exception("PDO_SMTP_AddIn: The sender-address must not be empty. Please check.");
                
            }

            mail.From = new MailAddress(SetEmailFrom);

            //mail.To.Add (SetEmailTo);
            //mail.CC.Add(Mail_CC);
            mail.Subject = SetSubject;
            mail.Body = Sethtmlbody;
            mail.IsBodyHtml = true;


            // SMTP-Client-object create and configure
            if (string.IsNullOrEmpty(SetsmtpServer))
            {
                throw new Exception("PDO_SMTP_AddIn: The smtpServer must not be empty. Please check.");

            }

            SmtpClient smtpClient = new SmtpClient(SetsmtpServer);
            smtpClient.Port =  ( SetIsSSL ) ? 587 : 465 ;
            smtpClient.EnableSsl = SetIsSSL; // Activate SSL, if necessary
            smtpClient.Credentials = new NetworkCredential(SetEmailFrom, SetPassword);
            

            // Expect100Continue = true needed for some provider, like strato.de
            ServicePointManager.Expect100Continue = true;  // if you have your own server, and want it faster: false; Might try with your provider.

            // TLS 1.2 as Security protocol, might change in future
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;


            // Send E-Mail , when nothing went wrong so far
            if (strResult == "")
            {
                try
                {
                    smtpClient.Send(mail);
                    Console.WriteLine("Yes, we did it!");
                    strResult = "Ok";
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Sorry, we didn't make it: " + ex.Message);
                    strResult = "Sorry, we didn't make it: " + ex.Message;
                }

            }

            // CleanUp
            mail.Dispose();
            smtpClient.Dispose();

            return strResult;


            // ****************************************
            // C|:-)  www.pdo.digital says Thank you !
            // ****************************************

        }



        static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }
}


