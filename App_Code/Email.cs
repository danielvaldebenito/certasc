using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Net.Mime;
/// <summary>
/// Descripción breve de Email
/// </summary>
public class Email
{
    public NetworkCredential Credentials { get; set; }
    public Email(string type)
    {
        Credentials = new NetworkCredential("dani270486@gmail.com", "commzgate");
        Send("Test", "daniel.valdebenito@commzgate-la.com");
    }

    public string GetHTML ()
    {
        var html = "<h1>Hola Mundo</h1>";
        return html;
    }
    public bool Send (string subject, string to) 
    {
        SmtpClient smtpServer = new SmtpClient("smtp.gmail.com");
        smtpServer.Port = 587;
        smtpServer.Credentials = Credentials;
        MailMessage mail = new MailMessage();
        mail.From = new MailAddress(Credentials.UserName);
        mail.To.Add(to); // para quien
        mail.CC.Add("dani270486@gmail.com"); // con copia
        mail.CC.Add(Credentials.UserName);
        mail.Subject = subject;
        mail.Body = GetHTML();
        mail.IsBodyHtml = true;
        ContentType contentType = new ContentType();
        contentType.MediaType = MediaTypeNames.Application.Octet;
        contentType.Name = "test.cs";
        mail.Attachments.Add(new Attachment("CreatePDF4401.cs", contentType));
        smtpServer.EnableSsl = true;
        smtpServer.Send(mail);
        return true;
    }
}