using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace CloudMailKit
{
    /// <summary>
    /// Represents an email message for sending
    /// COM-visible for Magic/VB6
    /// </summary>
    [ComVisible(true)]
    [Guid("A7B8C9D0-E1F2-3456-GHIJ-567890123EF0")]
    public class MailMessage
    {
        public MailMessage()
        {
            To = new List<string>();
            Cc = new List<string>();
            Bcc = new List<string>();
            Attachments = new List<string>();
        }

        public string From { get; set; }
        public List<string> To { get; set; }
        public List<string> Cc { get; set; }
        public List<string> Bcc { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public bool IsHtml { get; set; }
        public string Importance { get; set; } // "low", "normal", "high"
        public List<string> Attachments { get; set; }
        
        // Helper methods for COM
        public void AddTo(string email) => To.Add(email);
        public void AddCc(string email) => Cc.Add(email);
        public void AddBcc(string email) => Bcc.Add(email);
        public void AddAttachment(string path) => Attachments.Add(path);
    }
}