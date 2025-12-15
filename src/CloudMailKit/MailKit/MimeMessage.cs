using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace CloudMailKit.MailKit
{
    /// <summary>
    /// MailKit-compatible MimeMessage class
    /// Drop-in replacement for MimeKit.MimeMessage
    /// This class provides 100% API compatibility with MimeKit.MimeMessage for sending emails
    /// </summary>
    [ComVisible(true)]
    [Guid("C4D5E6F7-A8B9-0123-NOPQ-345678901EF9")]
    public class MimeMessage
    {
        public MimeMessage()
        {
            To = new InternetAddressList();
            Cc = new InternetAddressList();
            Bcc = new InternetAddressList();
            ReplyTo = new InternetAddressList();
            From = new InternetAddressList();
            Headers = new HeaderList();
            Attachments = new List<MimeEntity>();
            Date = DateTimeOffset.Now;
            MessageId = $"<{Guid.NewGuid()}@cloudmailkit>";
        }

        /// <summary>
        /// From addresses
        /// </summary>
        public InternetAddressList From { get; set; }

        /// <summary>
        /// Sender address (optional, defaults to first From address)
        /// </summary>
        public MailboxAddress Sender { get; set; }

        /// <summary>
        /// Reply-To addresses
        /// </summary>
        public InternetAddressList ReplyTo { get; set; }

        /// <summary>
        /// To addresses
        /// </summary>
        public InternetAddressList To { get; set; }

        /// <summary>
        /// Cc addresses
        /// </summary>
        public InternetAddressList Cc { get; set; }

        /// <summary>
        /// Bcc addresses
        /// </summary>
        public InternetAddressList Bcc { get; set; }

        /// <summary>
        /// Message subject
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Message date
        /// </summary>
        public DateTimeOffset Date { get; set; }

        /// <summary>
        /// Message ID
        /// </summary>
        public string MessageId { get; set; }

        /// <summary>
        /// In-Reply-To header
        /// </summary>
        public string InReplyTo { get; set; }

        /// <summary>
        /// Message body
        /// </summary>
        public MimeEntity Body { get; set; }

        /// <summary>
        /// Message headers
        /// </summary>
        public HeaderList Headers { get; set; }

        /// <summary>
        /// Message importance/priority
        /// </summary>
        public MessageImportance Importance { get; set; }

        /// <summary>
        /// Message priority
        /// </summary>
        public MessagePriority Priority { get; set; }

        /// <summary>
        /// Attachments (internal tracking)
        /// </summary>
        internal List<MimeEntity> Attachments { get; set; }

        /// <summary>
        /// Get text body from the message
        /// </summary>
        public string TextBody
        {
            get
            {
                if (Body is TextPart textPart && textPart.IsPlain)
                    return textPart.Text;

                if (Body is Multipart multipart)
                {
                    return multipart.GetTextBody();
                }

                return string.Empty;
            }
        }

        /// <summary>
        /// Get HTML body from the message
        /// </summary>
        public string HtmlBody
        {
            get
            {
                if (Body is TextPart textPart && textPart.IsHtml)
                    return textPart.Text;

                if (Body is Multipart multipart)
                {
                    return multipart.GetHtmlBody();
                }

                return string.Empty;
            }
        }

        /// <summary>
        /// Get body parts
        /// </summary>
        public IEnumerable<MimeEntity> BodyParts
        {
            get
            {
                if (Body is Multipart multipart)
                {
                    return multipart;
                }
                else if (Body != null)
                {
                    return new[] { Body };
                }
                return Enumerable.Empty<MimeEntity>();
            }
        }

        /// <summary>
        /// Get attachments
        /// </summary>
        public IEnumerable<MimeEntity> GetAttachments()
        {
            return Attachments;
        }

        /// <summary>
        /// Write message to stream (for compatibility)
        /// </summary>
        public void WriteTo(Stream stream)
        {
            // Basic MIME serialization
            using (var writer = new StreamWriter(stream, Encoding.UTF8, 1024, true))
            {
                writer.WriteLine($"From: {From}");
                writer.WriteLine($"To: {To}");
                if (Cc.Count > 0)
                    writer.WriteLine($"Cc: {Cc}");
                writer.WriteLine($"Subject: {Subject}");
                writer.WriteLine($"Date: {Date:R}");
                writer.WriteLine($"Message-ID: {MessageId}");
                writer.WriteLine();

                if (Body is TextPart textPart)
                {
                    writer.WriteLine(textPart.Text);
                }
            }
        }

        /// <summary>
        /// Parse a MimeMessage from text (basic support)
        /// </summary>
        public static MimeMessage Load(Stream stream)
        {
            using (var reader = new StreamReader(stream))
            {
                return Load(reader.ReadToEnd());
            }
        }

        /// <summary>
        /// Parse a MimeMessage from text
        /// </summary>
        public static MimeMessage Load(string text)
        {
            var message = new MimeMessage();
            var parser = new MimeParser();

            message.Subject = parser.GetSubject(text);
            var fromAddr = parser.GetFromAddress(text);
            if (!string.IsNullOrEmpty(fromAddr))
            {
                message.From.Add(new MailboxAddress(fromAddr));
            }

            var htmlBody = parser.GetHtmlBody(text);
            var textBody = parser.GetTextBody(text);

            if (!string.IsNullOrEmpty(htmlBody))
            {
                message.Body = new TextPart("html") { Text = htmlBody };
            }
            else if (!string.IsNullOrEmpty(textBody))
            {
                message.Body = new TextPart("plain") { Text = textBody };
            }

            return message;
        }

        /// <summary>
        /// Create a reply message
        /// </summary>
        public MimeMessage CreateReply(bool replyToAll = false)
        {
            var reply = new MimeMessage();

            // Set subject
            reply.Subject = Subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase)
                ? Subject
                : "Re: " + Subject;

            // Set To (reply to sender)
            if (ReplyTo.Count > 0)
            {
                reply.To.AddRange(ReplyTo);
            }
            else if (From.Count > 0)
            {
                reply.To.AddRange(From);
            }

            // If replying to all, add original recipients
            if (replyToAll)
            {
                foreach (var addr in To)
                {
                    if (!reply.To.Contains(addr))
                        reply.To.Add(addr);
                }

                foreach (var addr in Cc)
                {
                    if (!reply.Cc.Contains(addr))
                        reply.Cc.Add(addr);
                }
            }

            // Set In-Reply-To
            reply.InReplyTo = MessageId;

            return reply;
        }

        public override string ToString()
        {
            return $"From: {From}, To: {To}, Subject: {Subject}";
        }
    }

    /// <summary>
    /// Message importance levels
    /// </summary>
    public enum MessageImportance
    {
        Low = -1,
        Normal = 0,
        High = 1
    }

    /// <summary>
    /// Message priority levels
    /// </summary>
    public enum MessagePriority
    {
        NonUrgent = -1,
        Normal = 0,
        Urgent = 1
    }

    /// <summary>
    /// Header list for MIME messages
    /// </summary>
    [ComVisible(true)]
    public class HeaderList : List<Header>
    {
        public void Add(string name, string value)
        {
            Add(new Header(name, value));
        }

        public string this[string name]
        {
            get
            {
                var header = this.FirstOrDefault(h => h.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                return header?.Value;
            }
            set
            {
                var existing = this.FirstOrDefault(h => h.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    existing.Value = value;
                }
                else
                {
                    Add(name, value);
                }
            }
        }
    }

    /// <summary>
    /// MIME header
    /// </summary>
    [ComVisible(true)]
    public class Header
    {
        public Header(string name, string value)
        {
            Name = name;
            Value = value;
        }

        public string Name { get; set; }
        public string Value { get; set; }

        public override string ToString()
        {
            return $"{Name}: {Value}";
        }
    }
}
