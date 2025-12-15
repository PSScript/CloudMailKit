using System;
using System.Runtime.InteropServices;

namespace CloudMailKit
{
    /// <summary>
    /// Represents a mail folder from Graph API
    /// </summary>
    [ComVisible(true)]
    [Guid("D1E2F3A4-B5C6-7890-IJKL-890123456EF2")]
    public class GraphFolder
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string ParentFolderId { get; set; }
        public int ChildFolderCount { get; set; }
        public int UnreadItemCount { get; set; }
        public int TotalItemCount { get; set; }
    }
}
