using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace CloudMailKit.MailKit
{
    /// <summary>
    /// MailKit-compatible InternetAddress class
    /// Drop-in replacement for MimeKit.InternetAddress
    /// </summary>
    [ComVisible(true)]
    [Guid("F1A2B3C4-D5E6-7890-KLMN-012345678EF6")]
    public abstract class InternetAddress
    {
        public string Name { get; set; }

        public abstract string Address { get; }

        public override string ToString()
        {
            if (!string.IsNullOrEmpty(Name))
            {
                return $"{Name} <{Address}>";
            }
            return Address;
        }

        public static InternetAddress Parse(string text)
        {
            if (string.IsNullOrEmpty(text))
                throw new ArgumentNullException(nameof(text));

            return MailboxAddress.Parse(text);
        }

        public static bool TryParse(string text, out InternetAddress address)
        {
            try
            {
                address = Parse(text);
                return true;
            }
            catch
            {
                address = null;
                return false;
            }
        }
    }

    /// <summary>
    /// MailKit-compatible MailboxAddress class
    /// Drop-in replacement for MimeKit.MailboxAddress
    /// </summary>
    [ComVisible(true)]
    [Guid("A2B3C4D5-E6F7-8901-LMNO-123456789EF7")]
    public class MailboxAddress : InternetAddress
    {
        private string _address;

        public MailboxAddress(string name, string address)
        {
            Name = name;
            _address = address;
        }

        public MailboxAddress(string address)
        {
            Name = string.Empty;
            _address = address;
        }

        public override string Address => _address;

        public new static MailboxAddress Parse(string text)
        {
            if (string.IsNullOrEmpty(text))
                throw new ArgumentNullException(nameof(text));

            text = text.Trim();

            // Handle "Name <email@domain.com>" format
            var match = System.Text.RegularExpressions.Regex.Match(text, @"^(.*?)\s*<([^>]+)>$");
            if (match.Success)
            {
                var name = match.Groups[1].Value.Trim().Trim('"');
                var email = match.Groups[2].Value.Trim();
                return new MailboxAddress(name, email);
            }

            // Handle plain "email@domain.com" format
            return new MailboxAddress(string.Empty, text);
        }

        public new static bool TryParse(string text, out MailboxAddress address)
        {
            try
            {
                address = Parse(text);
                return true;
            }
            catch
            {
                address = null;
                return false;
            }
        }
    }

    /// <summary>
    /// MailKit-compatible InternetAddressList class
    /// Drop-in replacement for MimeKit.InternetAddressList
    /// </summary>
    [ComVisible(true)]
    [Guid("B3C4D5E6-F7A8-9012-MNOP-234567890EF8")]
    public class InternetAddressList : IList<InternetAddress>
    {
        private readonly List<InternetAddress> _addresses = new List<InternetAddress>();

        public InternetAddressList()
        {
        }

        public InternetAddressList(IEnumerable<InternetAddress> addresses)
        {
            if (addresses != null)
            {
                _addresses.AddRange(addresses);
            }
        }

        public int Count => _addresses.Count;

        public bool IsReadOnly => false;

        public InternetAddress this[int index]
        {
            get => _addresses[index];
            set => _addresses[index] = value;
        }

        public void Add(InternetAddress item)
        {
            if (item != null)
                _addresses.Add(item);
        }

        public void Add(string address)
        {
            if (!string.IsNullOrEmpty(address))
                _addresses.Add(MailboxAddress.Parse(address));
        }

        public void AddRange(IEnumerable<InternetAddress> addresses)
        {
            if (addresses != null)
                _addresses.AddRange(addresses);
        }

        public void Clear()
        {
            _addresses.Clear();
        }

        public bool Contains(InternetAddress item)
        {
            return _addresses.Contains(item);
        }

        public void CopyTo(InternetAddress[] array, int arrayIndex)
        {
            _addresses.CopyTo(array, arrayIndex);
        }

        public IEnumerator<InternetAddress> GetEnumerator()
        {
            return _addresses.GetEnumerator();
        }

        public int IndexOf(InternetAddress item)
        {
            return _addresses.IndexOf(item);
        }

        public void Insert(int index, InternetAddress item)
        {
            _addresses.Insert(index, item);
        }

        public bool Remove(InternetAddress item)
        {
            return _addresses.Remove(item);
        }

        public void RemoveAt(int index)
        {
            _addresses.RemoveAt(index);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public override string ToString()
        {
            return string.Join(", ", _addresses);
        }
    }
}
