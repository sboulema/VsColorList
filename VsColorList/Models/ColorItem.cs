using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace VsColorList.Models
{
    public class ColorItem
    {
        public string Name = string.Empty;

        public ThemeResourceKey Key;

        public ThemeResourceKeyType KeyType;

        public Guid Category;

        public Dictionary<string, Color> Colors = new Dictionary<string, Color>();
    }
}
