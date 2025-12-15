using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace CloudMailKit.MailKit
{
    /// <summary>
    /// Body builder for constructing MIME message bodies
    /// Drop-in replacement for MimeKit.BodyBuilder
    /// </summary>
    [ComVisible(true)]
    [Guid("B9C0D1E2-F3A4-5678-STUV-890123456EFE")]
    public class BodyBuilder
    {
        private readonly List<MimePart> _attachments = new List<MimePart>();
        private readonly List<MimePart> _linkedResources = new List<MimePart>();

        public BodyBuilder()
        {
        }

        /// <summary>
        /// Plain text body
        /// </summary>
        public string TextBody { get; set; }

        /// <summary>
        /// HTML body
        /// </summary>
        public string HtmlBody { get; set; }

        /// <summary>
        /// Attachments collection
        /// </summary>
        public IList<MimePart> Attachments => _attachments;

        /// <summary>
        /// Linked resources (for embedded images in HTML)
        /// </summary>
        public IList<MimePart> LinkedResources => _linkedResources;

        /// <summary>
        /// Build the message body
        /// </summary>
        /// <returns>The constructed MIME entity</returns>
        public MimeEntity ToMessageBody()
        {
            MimeEntity body = null;

            // Create text and/or HTML parts
            if (!string.IsNullOrEmpty(TextBody) && !string.IsNullOrEmpty(HtmlBody))
            {
                // Both text and HTML - create multipart/alternative
                var alternative = new Multipart("alternative");
                alternative.Add(new TextPart("plain") { Text = TextBody });

                if (_linkedResources.Count > 0)
                {
                    // HTML with linked resources - create multipart/related
                    var related = new Multipart("related");
                    related.Add(new TextPart("html") { Text = HtmlBody });
                    foreach (var resource in _linkedResources)
                    {
                        related.Add(resource);
                    }
                    alternative.Add(related);
                }
                else
                {
                    alternative.Add(new TextPart("html") { Text = HtmlBody });
                }

                body = alternative;
            }
            else if (!string.IsNullOrEmpty(HtmlBody))
            {
                // HTML only
                if (_linkedResources.Count > 0)
                {
                    var related = new Multipart("related");
                    related.Add(new TextPart("html") { Text = HtmlBody });
                    foreach (var resource in _linkedResources)
                    {
                        related.Add(resource);
                    }
                    body = related;
                }
                else
                {
                    body = new TextPart("html") { Text = HtmlBody };
                }
            }
            else if (!string.IsNullOrEmpty(TextBody))
            {
                // Text only
                body = new TextPart("plain") { Text = TextBody };
            }

            // Add attachments if present
            if (_attachments.Count > 0)
            {
                var mixed = new Multipart("mixed");

                if (body != null)
                {
                    mixed.Add(body);
                }

                foreach (var attachment in _attachments)
                {
                    mixed.Add(attachment);
                }

                return mixed;
            }

            return body ?? new TextPart("plain") { Text = string.Empty };
        }

        /// <summary>
        /// Add attachment from file path
        /// </summary>
        public void AddAttachment(string filePath)
        {
            var attachment = MimePart.CreateAttachment(filePath);
            _attachments.Add(attachment);
        }

        /// <summary>
        /// Add attachment from file path with content type
        /// </summary>
        public void AddAttachment(string filePath, string contentType)
        {
            var attachment = MimePart.CreateAttachment(filePath, contentType);
            _attachments.Add(attachment);
        }

        /// <summary>
        /// Add attachment from byte array
        /// </summary>
        public void AddAttachment(string fileName, byte[] data, string contentType = null)
        {
            if (string.IsNullOrEmpty(contentType))
            {
                contentType = "application/octet-stream";
            }

            var attachment = new MimePart(contentType)
            {
                FileName = fileName,
                Content = data,
                ContentDisposition = $"attachment; filename=\"{fileName}\"",
                ContentTransferEncoding = "base64"
            };

            _attachments.Add(attachment);
        }

        /// <summary>
        /// Add linked resource (embedded image)
        /// </summary>
        public MimePart AddLinkedResource(string filePath)
        {
            var resource = MimePart.CreateAttachment(filePath);
            resource.ContentId = $"<{Guid.NewGuid()}>";
            resource.ContentDisposition = "inline";
            _linkedResources.Add(resource);
            return resource;
        }
    }
}
