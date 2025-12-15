using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace CloudMailKit.MailKit
{
    /// <summary>
    /// Base class for MIME entities
    /// Drop-in replacement for MimeKit.MimeEntity
    /// </summary>
    [ComVisible(true)]
    [Guid("D5E6F7A8-B9C0-1234-OPQR-456789012EFA")]
    public abstract class MimeEntity
    {
        public MimeEntity()
        {
            Headers = new HeaderList();
            ContentType = new ContentType();
        }

        public HeaderList Headers { get; set; }
        public ContentType ContentType { get; set; }
        public string ContentId { get; set; }
        public string ContentDisposition { get; set; }
        public string ContentTransferEncoding { get; set; }

        public bool IsAttachment => ContentDisposition?.Contains("attachment") ?? false;
    }

    /// <summary>
    /// Text part for MIME messages
    /// Drop-in replacement for MimeKit.TextPart
    /// </summary>
    [ComVisible(true)]
    [Guid("E6F7A8B9-C0D1-2345-PQRS-567890123EFB")]
    public class TextPart : MimeEntity
    {
        public TextPart() : base()
        {
            ContentType = new ContentType("text", "plain");
        }

        public TextPart(string subtype) : base()
        {
            ContentType = new ContentType("text", subtype);
        }

        public string Text { get; set; }

        public bool IsPlain => ContentType.MimeType.Equals("text/plain", StringComparison.OrdinalIgnoreCase);
        public bool IsHtml => ContentType.MimeType.Equals("text/html", StringComparison.OrdinalIgnoreCase);

        public override string ToString()
        {
            return Text ?? string.Empty;
        }
    }

    /// <summary>
    /// Multipart MIME entity
    /// Drop-in replacement for MimeKit.Multipart
    /// </summary>
    [ComVisible(true)]
    [Guid("F7A8B9C0-D1E2-3456-QRST-678901234EFC")]
    public class Multipart : MimeEntity, IList<MimeEntity>
    {
        private readonly List<MimeEntity> _parts = new List<MimeEntity>();

        public Multipart() : base()
        {
            ContentType = new ContentType("multipart", "mixed");
        }

        public Multipart(string subtype) : base()
        {
            ContentType = new ContentType("multipart", subtype);
        }

        public int Count => _parts.Count;
        public bool IsReadOnly => false;

        public MimeEntity this[int index]
        {
            get => _parts[index];
            set => _parts[index] = value;
        }

        public void Add(MimeEntity item)
        {
            _parts.Add(item);
        }

        public void Clear()
        {
            _parts.Clear();
        }

        public bool Contains(MimeEntity item)
        {
            return _parts.Contains(item);
        }

        public void CopyTo(MimeEntity[] array, int arrayIndex)
        {
            _parts.CopyTo(array, arrayIndex);
        }

        public IEnumerator<MimeEntity> GetEnumerator()
        {
            return _parts.GetEnumerator();
        }

        public int IndexOf(MimeEntity item)
        {
            return _parts.IndexOf(item);
        }

        public void Insert(int index, MimeEntity item)
        {
            _parts.Insert(index, item);
        }

        public bool Remove(MimeEntity item)
        {
            return _parts.Remove(item);
        }

        public void RemoveAt(int index)
        {
            _parts.RemoveAt(index);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal string GetTextBody()
        {
            foreach (var part in _parts)
            {
                if (part is TextPart textPart && textPart.IsPlain)
                    return textPart.Text;

                if (part is Multipart multipart)
                {
                    var text = multipart.GetTextBody();
                    if (!string.IsNullOrEmpty(text))
                        return text;
                }
            }
            return string.Empty;
        }

        internal string GetHtmlBody()
        {
            foreach (var part in _parts)
            {
                if (part is TextPart textPart && textPart.IsHtml)
                    return textPart.Text;

                if (part is Multipart multipart)
                {
                    var html = multipart.GetHtmlBody();
                    if (!string.IsNullOrEmpty(html))
                        return html;
                }
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// MIME part for file attachments
    /// Drop-in replacement for MimeKit.MimePart
    /// </summary>
    [ComVisible(true)]
    [Guid("A8B9C0D1-E2F3-4567-RSTU-789012345EFD")]
    public class MimePart : MimeEntity
    {
        public MimePart() : base()
        {
        }

        public MimePart(string mimeType) : base()
        {
            var parts = mimeType.Split('/');
            if (parts.Length == 2)
            {
                ContentType = new ContentType(parts[0], parts[1]);
            }
            else
            {
                ContentType = new ContentType("application", "octet-stream");
            }
        }

        public string FileName { get; set; }
        public byte[] Content { get; set; }

        /// <summary>
        /// Create attachment from file
        /// </summary>
        public static MimePart CreateAttachment(string filePath, string contentType = null)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            var fileName = Path.GetFileName(filePath);
            var content = File.ReadAllBytes(filePath);

            if (string.IsNullOrEmpty(contentType))
            {
                contentType = GetMimeType(fileName);
            }

            var part = new MimePart(contentType)
            {
                FileName = fileName,
                Content = content,
                ContentDisposition = $"attachment; filename=\"{fileName}\"",
                ContentTransferEncoding = "base64"
            };

            return part;
        }

        private static string GetMimeType(string fileName)
        {
            var ext = Path.GetExtension(fileName).ToLowerInvariant();

            var mimeTypes = new Dictionary<string, string>
            {
                { ".txt", "text/plain" },
                { ".pdf", "application/pdf" },
                { ".doc", "application/msword" },
                { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                { ".xls", "application/vnd.ms-excel" },
                { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                { ".png", "image/png" },
                { ".jpg", "image/jpeg" },
                { ".jpeg", "image/jpeg" },
                { ".gif", "image/gif" },
                { ".zip", "application/zip" },
                { ".csv", "text/csv" },
                { ".xml", "application/xml" },
                { ".json", "application/json" }
            };

            return mimeTypes.TryGetValue(ext, out var mimeType) ? mimeType : "application/octet-stream";
        }
    }

    /// <summary>
    /// Content type helper class
    /// </summary>
    [ComVisible(true)]
    public class ContentType
    {
        public ContentType()
        {
            MediaType = "text";
            MediaSubtype = "plain";
        }

        public ContentType(string mediaType, string mediaSubtype)
        {
            MediaType = mediaType;
            MediaSubtype = mediaSubtype;
        }

        public string MediaType { get; set; }
        public string MediaSubtype { get; set; }
        public string Charset { get; set; } = "utf-8";
        public string Name { get; set; }
        public string Boundary { get; set; }

        public string MimeType => $"{MediaType}/{MediaSubtype}";

        public override string ToString()
        {
            return MimeType;
        }
    }
}
