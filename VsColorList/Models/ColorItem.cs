using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace VsColorList.Models
{
    public class ColorItem
    {
        public string Name { get; set; } = string.Empty;

        public ThemeResourceKey Key { get; set; }

        public ThemeResourceKeyType? KeyType { get; set; }

        public Guid? Category { get; set; }

        public Dictionary<string, Color> Colors { get; set; } = new Dictionary<string, Color>();

        public string Type { get; set; } = string.Empty;
    }
}
