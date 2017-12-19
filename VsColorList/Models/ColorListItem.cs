using Microsoft.VisualStudio.Shell;
using System;
using System.Drawing;

namespace VsColorList.Models
{
    public class ColorListItem
    {
        public string Name;
        public ThemeResourceKey Key;
        public ThemeResourceKeyType KeyType;
        public Guid Category;
        public Color LightColor;
        public Color DarkColor;
        public Color BlueColor;
    }
}
