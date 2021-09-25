using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using System.Collections.Generic;
using System.Linq;
using VsColorList.Models;

namespace VsColorList.Helpers
{
    public static class ColorListHelper
    {
        public static List<ColorItem> GetColorList(string themeName)
        {
            var colorList = new List<ColorItem>();

            foreach (var color in typeof(EnvironmentColors).GetProperties())
            {
                if (!(color.GetValue(null) is ThemeResourceKey themeResourceKey))
                {
                    continue;
                }

                var colorItem = new ColorItem
                {
                    Key = themeResourceKey,
                    Colors = { { themeName, VSColorTheme.GetThemedColor(themeResourceKey) } }
                };

                colorList.Add(colorItem);
            }

            return colorList;
        }

        public static List<ColorItem> GetEnvironmentColorList(List<ColorItem> lightColors,
            List<ColorItem> darkColors, List<ColorItem> blueColors)
        {
            var colorList = new List<ColorItem>();

            for (var i = 0; i < lightColors.Count; i++)
            {
                var colorItem = new ColorItem
                {
                    Key = lightColors[i].Key,
                    KeyType = lightColors[i].Key.KeyType,
                    Category = lightColors[i].Key.Category
                };

                colorItem.Colors = lightColors[i].Colors
                    .Union(darkColors[i].Colors)
                    .Union(blueColors[i].Colors)
                    .ToDictionary(pair => pair.Key, pair => pair.Value);

                colorList.Add(colorItem);
            }

            return colorList;
        }

        public static List<ColorItem> GetVsColorList(List<KeyValuePair<ThemeResourceKey, uint>> lightColors,
            List<KeyValuePair<ThemeResourceKey, uint>> darkColors, List<KeyValuePair<ThemeResourceKey, uint>> blueColors)
        {
            var colorList = new List<ColorItem>();

            for (var i = 0; i < lightColors.Count; i++)
            {
                colorList.Add(new ColorItem
                {
                    Key = lightColors[i].Key,
                    KeyType = lightColors[i].Key.KeyType,
                    Category = lightColors[i].Key.Category,
                    Colors = {
                        { "light", VSColorTheme.GetThemedColor(lightColors[i].Key) },
                        { "dark", VSColorTheme.GetThemedColor(darkColors[i].Key) },
                        { "blue", VSColorTheme.GetThemedColor(blueColors[i].Key) },
                    }
                });
            }

            return colorList;
        }
    }
}
