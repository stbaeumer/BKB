using System.Net;
using System.Net.Mail;
using MailKit.Net.Smtp;
using MimeKit;

public class Mail
{
    public Mail() {}

    public void Senden(string subject, string absendername, string body, string attachment, string receiverEmail)
    {
        try{
            string smtpServer = Global.SmtpServer;
            int smtpPort = Convert.ToInt32(Global.SmtpPort);
            string senderEmail = Global.SmtpUser;

            if(Global.SmtpPassword == null || Global.SmtpPassword.Length <= 3)
            {
                Console.WriteLine("Bitte geben Sie das Passwort von " + Global.SmtpUser +" für den E-Mail-Versand ein:");
                Global.SmtpPassword = Console.ReadLine();
            }
            
            string senderPassword = Global.SmtpPassword;

            var email = new MimeMessage();
            email.From.Add(new MailboxAddress(Global.SmtpUser, senderEmail));
            email.To.Add(new MailboxAddress("Empfänger", receiverEmail));
            //email.To.Add(new MailboxAddress("Empfänger", "stefan.baeumer@berufskolleg-borken.de"));
            
            email.Subject = subject;
            email.Cc.Add(new MailboxAddress("Empfänger", "stefan.baeumer@berufskolleg-borken.de"));
            email.Bcc.Add(new MailboxAddress("Empfänger", "catrin.stakenkoetter@berufskolleg-borken.de"));

            // 1️⃣ Erstelle den Haupttext der E-Mail
            var textPart = new TextPart("plain") { Text = body };

            // 2️⃣ Falls eine Datei angegeben wurde, erstelle den Anhang
            var multipart = new Multipart("mixed");
            multipart.Add(textPart); // Erst den Text hinzufügen

            if (!string.IsNullOrEmpty(attachment) && System.IO.File.Exists(attachment))
            {
                var attachmentPart = new MimePart()
                {
                    Content = new MimeContent(System.IO.File.OpenRead(attachment)),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    FileName = System.IO.Path.GetFileName(attachment)
                };

                multipart.Add(attachmentPart);
            }

            // 3️⃣ Setze den E-Mail-Body auf multipart (Text + Anhang)
            email.Body = multipart;

            using (var smtpClient = new MailKit.Net.Smtp.SmtpClient())
            {
                smtpClient.ServerCertificateValidationCallback = (s, c, h, e) => true; // SSL-Zertifikatsvalidierung deaktivieren
                smtpClient.Connect(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls);
                smtpClient.Authenticate(senderEmail, senderPassword);
                smtpClient.Send(email);
                smtpClient.Disconnect(true);
            }

            Global.ZeileSchreiben(receiverEmail, "gesendet", ConsoleColor.Green, ConsoleColor.White);        
        }
        catch(Exception ex){
            Global.ZeileSchreiben(receiverEmail, "Versand gescheitert.", ConsoleColor.Red,ConsoleColor.Gray);        
            Console.WriteLine("Fehler beim Versand der E-Mail an " + receiverEmail + ": " + ex.Message);
        }
    }

    public void Senden(string subject, string sender, string body, Stream attachmentStream, string attachmentName, string receiver)
    {
        var mailMessage = new MailMessage(sender, receiver, subject, body);
        mailMessage.Attachments.Add(new Attachment(attachmentStream, attachmentName));

        if(Global.SmtpPassword == null || Global.SmtpPassword.Length <= 3)
        {
            Console.WriteLine("Bitte geben Sie das Passwort von " + Global.SmtpUser +" für den E-Mail-Versand ein:");
            Global.SmtpPassword = Console.ReadLine();
        }

        using (var smtpClient = new System.Net.Mail.SmtpClient(Global.SmtpServer, Convert.ToInt32(Global.SmtpPort)))
        {
            smtpClient.Credentials = new NetworkCredential(Global.SmtpUser, Global.SmtpPassword);
            smtpClient.EnableSsl = true;
            smtpClient.Send(mailMessage);
        }
    }
}