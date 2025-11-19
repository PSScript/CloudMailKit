using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace CloudMailKit
{
    /// <summary>
    /// Unified mail client combining reader and sender capabilities
    /// COM-visible for Magic/VB6 integration
    /// </summary>
    [ComVisible(true)]
    [Guid("E5F6A7B8-C9D0-1234-EFGH-345678901CDE")]
    [ClassInterface(ClassInterfaceType.None)]
    public class UnifiedMailClient : IMailClient, IDisposable
    {
        private readonly GraphMailReader _reader;
        private readonly GraphMailSender _sender;
        private readonly MimeParser _parser;
        private bool _isInitialized = false;

        #region Constructor

        public UnifiedMailClient()
        {
            _reader = new GraphMailReader();
            _sender = new GraphMailSender();
            _parser = new MimeParser();
        }

        #endregion

        #region Initialization

        /// <summary>
        /// Initialize with OAuth credentials
        /// </summary>
        public void Initialize(string tenantId, string clientId, string clientSecret, string mailboxAddress)
        {
            _reader.Initialize(tenantId, clientId, clientSecret, mailboxAddress);
            _sender.Initialize(tenantId, clientId, clientSecret, mailboxAddress);
            _isInitialized = true;
        }

        /// <summary>
        /// Initialize from config file (app.config/web.config)
        /// </summary>
        public void InitializeFromConfig()
        {
            var tenantId = System.Configuration.ConfigurationManager.AppSettings["CloudMailKit.TenantId"];
            var clientId = System.Configuration.ConfigurationManager.AppSettings["CloudMailKit.ClientId"];
            var clientSecret = System.Configuration.ConfigurationManager.AppSettings["CloudMailKit.ClientSecret"];
            var mailbox = System.Configuration.ConfigurationManager.AppSettings["CloudMailKit.MailboxAddress"];

            if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) ||
                string.IsNullOrEmpty(clientSecret) || string.IsNullOrEmpty(mailbox))
            {
                throw new InvalidOperationException("Missing configuration. Required appSettings: CloudMailKit.TenantId, CloudMailKit.ClientId, CloudMailKit.ClientSecret, CloudMailKit.MailboxAddress");
            }

            Initialize(tenantId, clientId, clientSecret, mailbox);
        }

        private void EnsureInitialized()
        {
            if (!_isInitialized)
            {
                throw new InvalidOperationException("Not initialized. Call Initialize() first.");
            }
        }

        #endregion

        #region Reader Methods (Delegated)

        public GraphFolder GetInbox()
        {
            EnsureInitialized();
            return _reader.GetInbox();
        }

        public GraphFolder GetFolder(string folderName)
        {
            EnsureInitialized();
            return _reader.GetFolder(folderName);
        }

        public List<GraphFolder> ListFolders()
        {
            EnsureInitialized();
            return _reader.ListFolders();
        }

        public List<GraphMessage> ListMessages(string folderId, int maxCount = 100)
        {
            EnsureInitialized();
            return _reader.ListMessages(folderId, maxCount);
        }

        public string GetMessageMime(string messageId)
        {
            EnsureInitialized();
            return _reader.GetMessageMime(messageId);
        }

        public GraphMessage GetMessage(string messageId)
        {
            EnsureInitialized();
            return _reader.GetMessage(messageId);
        }

        public List<GraphMessage> SearchMessages(string query, string folderId = null)
        {
            EnsureInitialized();
            return _reader.SearchMessages(query, folderId);
        }

        public void MarkAsRead(string messageId)
        {
            EnsureInitialized();
            _reader.MarkAsRead(messageId);
        }

        public void MarkAsUnread(string messageId)
        {
            EnsureInitialized();
            _reader.MarkAsUnread(messageId);
        }

        public string MoveMessage(string messageId, string destinationFolderId)
        {
            EnsureInitialized();
            return _reader.MoveMessage(messageId, destinationFolderId);
        }

        public void DeleteMessage(string messageId)
        {
            EnsureInitialized();
            _reader.DeleteMessage(messageId);
        }

        #endregion

        #region Sender Methods (Delegated)

        public void SendMessage(MailMessage message)
        {
            EnsureInitialized();
            _sender.SendMessage(message);
        }

        public void SendSimple(string from, string to, string subject, string body, bool isHtml = false)
        {
            EnsureInitialized();
            _sender.SendSimple(from, to, subject, body, isHtml);
        }

        public void SendWithAttachment(string from, string to, string subject, string body, string attachmentPath)
        {
            EnsureInitialized();
            _sender.SendWithAttachment(from, to, subject, body, attachmentPath);
        }

        #endregion

        #region Parser Methods (Delegated)

        public string GetTextBody(string mimeContent)
        {
            return _parser.GetTextBody(mimeContent);
        }

        public string GetHtmlBody(string mimeContent)
        {
            return _parser.GetHtmlBody(mimeContent);
        }

        public string GetHeader(string mimeContent, string headerName)
        {
            return _parser.GetHeader(mimeContent, headerName);
        }

        public int GetAttachmentCount(string mimeContent)
        {
            return _parser.GetAttachmentCount(mimeContent);
        }

        public string GetAttachmentInfo(string mimeContent, int index)
        {
            return _parser.GetAttachmentInfo(mimeContent, index);
        }

        public void SaveAttachment(string mimeContent, int index, string outputPath)
        {
            _parser.SaveAttachment(mimeContent, index, outputPath);
        }

        public string GetSubject(string mimeContent)
        {
            return _parser.GetSubject(mimeContent);
        }

        public string GetFromAddress(string mimeContent)
        {
            return _parser.GetFromAddress(mimeContent);
        }

        #endregion

        #region High-Level Convenience Methods

        /// <summary>
        /// Get unread messages from inbox
        /// </summary>
        public List<GraphMessage> GetUnreadMessages(int maxCount = 50)
        {
            EnsureInitialized();
            var inbox = GetInbox();
            var allMessages = ListMessages(inbox.Id, maxCount);
            var unread = new List<GraphMessage>();

            foreach (var msg in allMessages)
            {
                if (!msg.IsRead)
                {
                    unread.Add(msg);
                }
            }

            return unread;
        }

        /// <summary>
        /// Reply to a message
        /// </summary>
        public void ReplyToMessage(string messageId, string replyBody, bool replyAll = false)
        {
            EnsureInitialized();

            // Get original message
            var mime = GetMessageMime(messageId);
            var originalFrom = GetFromAddress(mime);
            var originalSubject = GetSubject(mime);

            // Build reply subject
            var replySubject = originalSubject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase)
                ? originalSubject
                : "Re: " + originalSubject;

            // Get mailbox from reader
            var mailbox = _reader.GetMailboxAddress();

            // Send reply
            SendSimple(mailbox, originalFrom, replySubject, replyBody, false);
        }

        /// <summary>
        /// Forward a message
        /// </summary>
        public void ForwardMessage(string messageId, string toAddress, string additionalComments = "")
        {
            EnsureInitialized();

            var mime = GetMessageMime(messageId);
            var originalSubject = GetSubject(mime);
            var originalBody = GetTextBody(mime);

            var forwardSubject = originalSubject.StartsWith("Fwd:", StringComparison.OrdinalIgnoreCase)
                ? originalSubject
                : "Fwd: " + originalSubject;

            var forwardBody = additionalComments + "\n\n----- Forwarded Message -----\n\n" + originalBody;

            var mailbox = _reader.GetMailboxAddress();
            SendSimple(mailbox, toAddress, forwardSubject, forwardBody, false);
        }

        /// <summary>
        /// Process inbox with callback
        /// </summary>
        public void ProcessInbox(Action<GraphMessage, string> processCallback, int maxMessages = 100)
        {
            EnsureInitialized();

            var inbox = GetInbox();
            var messages = ListMessages(inbox.Id, maxMessages);

            foreach (var msg in messages)
            {
                var mime = GetMessageMime(msg.Id);
                processCallback(msg, mime);
            }
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            _reader?.Dispose();
            _sender?.Dispose();
        }

        #endregion
    }

    #region COM Interface

    [ComVisible(true)]
    [Guid("F6A7B8C9-D0E1-2345-FGHI-456789012DEF")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMailClient
    {
        // Initialization
        void Initialize(string tenantId, string clientId, string clientSecret, string mailboxAddress);
        void InitializeFromConfig();

        // Reader
        GraphFolder GetInbox();
        GraphFolder GetFolder(string folderName);
        List<GraphFolder> ListFolders();
        List<GraphMessage> ListMessages(string folderId, int maxCount);
        string GetMessageMime(string messageId);
        GraphMessage GetMessage(string messageId);
        List<GraphMessage> SearchMessages(string query, string folderId);
        void MarkAsRead(string messageId);
        void MarkAsUnread(string messageId);
        string MoveMessage(string messageId, string destinationFolderId);
        void DeleteMessage(string messageId);

        // Sender
        void SendMessage(MailMessage message);
        void SendSimple(string from, string to, string subject, string body, bool isHtml);
        void SendWithAttachment(string from, string to, string subject, string body, string attachmentPath);

        // Parser
        string GetTextBody(string mimeContent);
        string GetHtmlBody(string mimeContent);
        string GetHeader(string mimeContent, string headerName);
        int GetAttachmentCount(string mimeContent);
        string GetAttachmentInfo(string mimeContent, int index);
        void SaveAttachment(string mimeContent, int index, string outputPath);
        string GetSubject(string mimeContent);
        string GetFromAddress(string mimeContent);

        // High-level
        List<GraphMessage> GetUnreadMessages(int maxCount);
        void ReplyToMessage(string messageId, string replyBody, bool replyAll);
        void ForwardMessage(string messageId, string toAddress, string additionalComments);
    }

    #endregion
}