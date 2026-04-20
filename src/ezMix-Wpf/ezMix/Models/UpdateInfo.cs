using System.Collections.Generic;

namespace ezMix.Models
{
    public class UpdateInfo
    {
        public string Version { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
        public bool Mandatory { get; set; }
        public List<string> Changelog { get; set; } = new List<string>();
    }
}
