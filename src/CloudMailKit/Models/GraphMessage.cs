using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace CloudMailKit
{
    /// <summary>
    /// Represents an email message from Graph API
    /// </summary>
    [ComVisible(true)]
    [Guid("E2F3A4B5-C6D7-8901-JKLM-901234567EF3")]
    public class GraphMessage
    {
        public GraphMessage()
        {
            ToRecipients = new List<string>();
            CcRecipients = new List<string>();
            BccRecipients = new List<string>();
        }

        public string Id { get; set; }
        public string Subject { get; set; }
        public string BodyPreview { get; set; }
        public string BodyContent { get; set; }
        public string BodyContentType { get; set; } // "text" or "html"
        public string From { get; set; }
        public string Sender { get; set; }
        public List<string> ToRecipients { get; set; }
        public List<string> CcRecipients { get; set; }
        public List<string> BccRecipients { get; set; }
        public bool IsRead { get; set; }
        public bool IsDraft { get; set; }
        public string Importance { get; set; }
        public DateTime ReceivedDateTime { get; set; }
        public DateTime SentDateTime { get; set; }
        public bool HasAttachments { get; set; }
        public string InternetMessageId { get; set; }
        public string ConversationId { get; set; }
    }
}
