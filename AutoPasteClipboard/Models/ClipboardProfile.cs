using LiteDB;
using System.Collections.Generic;

namespace AutoPasteClipboard.Models
{
    public class ClipboardProfile
    {
        [BsonId]
        public string Profile { get; set; }
        public List<string> Clipboard { get; set; }

        public ClipboardProfile() { }
    }
}
