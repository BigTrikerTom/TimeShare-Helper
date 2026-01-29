# Mail.dll

**Mail.dll** is a robust, commercial-grade **.NET** library for working with **IMAP**, **SMTP**, **POP3**, and **MIME** protocols.

It includes a powerful email and MIME parser, enabling you to send, receive, and process email messages in .NET applications.

The library supports:
- **IMAP IDLE** (push email)
- **OAuth2 authentication** (Gmail, Outlook)
- **S/MIME encryption and signing**
- Advanced **MIME parsing**

Additional features include:
- Built-in **SSL/TLS** support
- **DKIM** signing and verification
- A secure **S/MIME** parser for encrypted and digitally signed messages
- Parsers for **iCal**, **vCard**, and Outlook's **.msg** file format
- Support for decoding **winmail.dat** attachments

## Getting started with Mail.dll

- [Mail.dll main website](https://www.limilabs.com/mail)
- [Sample library](https://www.limilabs.com/mail/samples)
- [Technical Q&A forum](https://www.limilabs.com/qa)

### Download emails using IMAP

```cs
using(Imap imap = new Imap())
{
    imap.ConnectSSL("imap.server.com");  // or Connect for non SSL/TLS
    imap.UseBestLogin("user", "password");

    imap.SelectInbox();
    List<long> uids = imap.Search(Flag.Unseen);
    foreach (long uid in uids)
    {
        var eml = imap.GetMessageByUID(uid);

        IMail email = new MailBuilder().CreateFromEml(eml);

        string subject = email.Subject;
    }
    imap.Close();
}
```

### Download emails using POP3

```cs
using(Pop3 pop3 = new Pop3())
{
    pop3.ConnectSSL("pop3.server.com");  // or Connect for non SSL/TLS   
    pop3.UseBestLogin("user", "password");

    List<string> uids = pop3.GetAll();
    foreach (string uid in uids)
    {
        var eml = pop3.GetMessageByUID(uid);
        IMail email = new MailBuilder().CreateFromEml(eml);

        string subject = email.Subject;
    }
    pop3.Close();
} 
```

### Send emails using SMTP

```cs
MailBuilder builder = new MailBuilder();
builder.From.Add(new MailBox("from@example.com"));
builder.To.Add(new MailBox("to@example.com"));
builder.Subject = "Subject";
builder.Html = @"Html with an image: <img src=""cid:lena"" />";

var visual = builder.AddVisual(@"c:\lena.jpeg");
visual.ContentId = "lena";

var attachment = builder.AddAttachment(@"c:\tmp.doc");
attachment.SetFileName("document.doc", guessContentType: true);

IMail email = builder.Create();

using(Smtp smtp = new Smtp())
{
    smtp.Connect("smtp.server.com");  // or ConnectSSL for SSL
    smtp.UseBestLogin("user", "password");

    smtp.SendMessage(email);                      
    smtp.Close();    
}
```


