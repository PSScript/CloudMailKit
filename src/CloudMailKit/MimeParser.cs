using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace CloudMailKit
{
    /// <summary>
    /// Simple MIME message parser
    /// Provides basic parsing without external dependencies
    /// </summary>
    [ComVisible(true)]
    [Guid("D9E0F1A2-B3C4-5678-IJKL-890123456EF5")]
    internal class MimeParser
    {
        public string GetTextBody(string mimeContent)
        {
            if (string.IsNullOrEmpty(mimeContent))
                return string.Empty;

            // Look for text/plain part
            var parts = SplitMimeParts(mimeContent);
            foreach (var part in parts)
            {
                if (part.ContentType.Contains("text/plain"))
                {
                    return DecodeBody(part.Body, part.TransferEncoding);
                }
            }

            return string.Empty;
        }

        public string GetHtmlBody(string mimeContent)
        {
            if (string.IsNullOrEmpty(mimeContent))
                return string.Empty;

            // Look for text/html part
            var parts = SplitMimeParts(mimeContent);
            foreach (var part in parts)
            {
                if (part.ContentType.Contains("text/html"))
                {
                    return DecodeBody(part.Body, part.TransferEncoding);
                }
            }

            return string.Empty;
        }

        public string GetHeader(string mimeContent, string headerName)
        {
            if (string.IsNullOrEmpty(mimeContent) || string.IsNullOrEmpty(headerName))
                return string.Empty;

            var lines = mimeContent.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var headerPrefix = headerName + ":";

            foreach (var line in lines)
            {
                if (line.StartsWith(headerPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    return line.Substring(headerPrefix.Length).Trim();
                }
            }

            return string.Empty;
        }

        public string GetSubject(string mimeContent)
        {
            return DecodeHeaderValue(GetHeader(mimeContent, "Subject"));
        }

        public string GetFromAddress(string mimeContent)
        {
            var from = GetHeader(mimeContent, "From");
            return ExtractEmailAddress(from);
        }

        public int GetAttachmentCount(string mimeContent)
        {
            if (string.IsNullOrEmpty(mimeContent))
                return 0;

            var parts = SplitMimeParts(mimeContent);
            return parts.Count(p => p.IsAttachment);
        }

        public string GetAttachmentInfo(string mimeContent, int index)
        {
            if (string.IsNullOrEmpty(mimeContent))
                return string.Empty;

            var parts = SplitMimeParts(mimeContent);
            var attachments = parts.Where(p => p.IsAttachment).ToList();

            if (index < 0 || index >= attachments.Count)
                return string.Empty;

            var attachment = attachments[index];
            return $"Filename: {attachment.Filename}, Size: {attachment.Body.Length} bytes, Type: {attachment.ContentType}";
        }

        public void SaveAttachment(string mimeContent, int index, string outputPath)
        {
            if (string.IsNullOrEmpty(mimeContent))
                throw new ArgumentException("MIME content cannot be empty");

            var parts = SplitMimeParts(mimeContent);
            var attachments = parts.Where(p => p.IsAttachment).ToList();

            if (index < 0 || index >= attachments.Count)
                throw new ArgumentOutOfRangeException(nameof(index), "Attachment index out of range");

            var attachment = attachments[index];
            var data = DecodeAttachment(attachment.Body, attachment.TransferEncoding);

            File.WriteAllBytes(outputPath, data);
        }

        private List<MimePart> SplitMimeParts(string mimeContent)
        {
            var parts = new List<MimePart>();

            // Find boundary
            var boundaryMatch = Regex.Match(mimeContent, @"boundary=""?([^""\s;]+)""?", RegexOptions.IgnoreCase);
            if (!boundaryMatch.Success)
            {
                // Single part message
                parts.Add(ParseSinglePart(mimeContent));
                return parts;
            }

            var boundary = "--" + boundaryMatch.Groups[1].Value;
            var sections = mimeContent.Split(new[] { boundary }, StringSplitOptions.None);

            foreach (var section in sections)
            {
                if (string.IsNullOrWhiteSpace(section) || section.Trim() == "--")
                    continue;

                parts.Add(ParseSinglePart(section));
            }

            return parts;
        }

        private MimePart ParseSinglePart(string content)
        {
            var part = new MimePart();
            var lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            var inHeaders = true;
            var bodyLines = new List<string>();

            foreach (var line in lines)
            {
                if (inHeaders)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        inHeaders = false;
                        continue;
                    }

                    if (line.StartsWith("Content-Type:", StringComparison.OrdinalIgnoreCase))
                    {
                        part.ContentType = line.Substring("Content-Type:".Length).Trim();
                    }
                    else if (line.StartsWith("Content-Transfer-Encoding:", StringComparison.OrdinalIgnoreCase))
                    {
                        part.TransferEncoding = line.Substring("Content-Transfer-Encoding:".Length).Trim();
                    }
                    else if (line.StartsWith("Content-Disposition:", StringComparison.OrdinalIgnoreCase))
                    {
                        var disposition = line.Substring("Content-Disposition:".Length).Trim();
                        if (disposition.Contains("attachment"))
                        {
                            part.IsAttachment = true;
                            var filenameMatch = Regex.Match(disposition, @"filename=""?([^""]+)""?", RegexOptions.IgnoreCase);
                            if (filenameMatch.Success)
                            {
                                part.Filename = filenameMatch.Groups[1].Value;
                            }
                        }
                    }
                }
                else
                {
                    bodyLines.Add(line);
                }
            }

            part.Body = string.Join("\n", bodyLines);
            return part;
        }

        private string DecodeBody(string body, string transferEncoding)
        {
            if (string.IsNullOrEmpty(body))
                return string.Empty;

            if (transferEncoding.Contains("base64", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var cleanedBody = body.Replace("\r", "").Replace("\n", "").Replace(" ", "");
                    var bytes = Convert.FromBase64String(cleanedBody);
                    return Encoding.UTF8.GetString(bytes);
                }
                catch
                {
                    return body;
                }
            }
            else if (transferEncoding.Contains("quoted-printable", StringComparison.OrdinalIgnoreCase))
            {
                return DecodeQuotedPrintable(body);
            }

            return body;
        }

        private byte[] DecodeAttachment(string body, string transferEncoding)
        {
            if (string.IsNullOrEmpty(body))
                return new byte[0];

            if (transferEncoding.Contains("base64", StringComparison.OrdinalIgnoreCase))
            {
                var cleanedBody = body.Replace("\r", "").Replace("\n", "").Replace(" ", "");
                return Convert.FromBase64String(cleanedBody);
            }

            return Encoding.UTF8.GetBytes(body);
        }

        private string DecodeQuotedPrintable(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            var output = new StringBuilder();
            var lines = input.Split('\n');

            foreach (var line in lines)
            {
                var trimmedLine = line.TrimEnd('\r');
                var i = 0;

                while (i < trimmedLine.Length)
                {
                    if (trimmedLine[i] == '=')
                    {
                        if (i + 2 < trimmedLine.Length)
                        {
                            var hex = trimmedLine.Substring(i + 1, 2);
                            try
                            {
                                var charCode = Convert.ToInt32(hex, 16);
                                output.Append((char)charCode);
                                i += 3;
                                continue;
                            }
                            catch
                            {
                                // Invalid hex, just append the character
                            }
                        }
                    }

                    output.Append(trimmedLine[i]);
                    i++;
                }

                output.Append("\n");
            }

            return output.ToString().TrimEnd('\n');
        }

        private string DecodeHeaderValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            // Handle RFC 2047 encoded words: =?charset?encoding?text?=
            var pattern = @"=\?([^?]+)\?([BQbq])\?([^?]+)\?=";
            var matches = Regex.Matches(value, pattern);

            if (matches.Count == 0)
                return value;

            var result = value;
            foreach (Match match in matches)
            {
                var charset = match.Groups[1].Value;
                var encoding = match.Groups[2].Value.ToUpper();
                var encodedText = match.Groups[3].Value;

                string decodedText;
                if (encoding == "B")
                {
                    // Base64
                    var bytes = Convert.FromBase64String(encodedText);
                    decodedText = Encoding.UTF8.GetString(bytes);
                }
                else if (encoding == "Q")
                {
                    // Quoted-printable
                    encodedText = encodedText.Replace('_', ' ');
                    decodedText = DecodeQuotedPrintable(encodedText);
                }
                else
                {
                    decodedText = encodedText;
                }

                result = result.Replace(match.Value, decodedText);
            }

            return result;
        }

        private string ExtractEmailAddress(string addressField)
        {
            if (string.IsNullOrEmpty(addressField))
                return string.Empty;

            // Extract email from formats like "Name <email@domain.com>" or just "email@domain.com"
            var match = Regex.Match(addressField, @"<([^>]+)>");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            // Try to find email pattern
            match = Regex.Match(addressField, @"[\w\.-]+@[\w\.-]+\.\w+");
            if (match.Success)
            {
                return match.Value;
            }

            return addressField.Trim();
        }

        private class MimePart
        {
            public string ContentType { get; set; } = string.Empty;
            public string TransferEncoding { get; set; } = "7bit";
            public string Body { get; set; } = string.Empty;
            public bool IsAttachment { get; set; } = false;
            public string Filename { get; set; } = string.Empty;
        }
    }
}
