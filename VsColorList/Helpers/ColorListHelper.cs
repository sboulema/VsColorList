using Community.VisualStudio.Toolkit;
using Microsoft.VisualStudio.Language.StandardClassification;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Classification;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Media;
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

        public static async Task<List<ColorItem>> GetClassificationColorList(string themeName)
        {
            var componentModel = await VS.Services.GetComponentModelAsync();
            var classificationFormatMapService = componentModel.GetService<IClassificationFormatMapService>();
            var classificationTypeRegistryService = componentModel.GetService<IClassificationTypeRegistryService>();

            var formatMap = classificationFormatMapService.GetClassificationFormatMap("text");

            var classificationTypeNames = typeof(PredefinedClassificationTypeNames)
                .GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                .Where(fi => fi.IsLiteral && !fi.IsInitOnly)
                .Select(fi => fi.GetRawConstantValue() as string)
                .ToList();

            var colorList = new List<ColorItem>();

            foreach (var name in classificationTypeNames)
            {
                var classificationType = classificationTypeRegistryService.GetClassificationType(name);

                if (classificationType == null)
                {
                    continue;
                }

                var textProperties = formatMap.GetTextProperties(classificationType);
                var foregroundBrush = textProperties.ForegroundBrush as SolidColorBrush;

                var colorItem = new ColorItem
                {
                    Name = classificationType.Classification,
                    Colors = { { themeName, System.Drawing.Color.FromArgb(
                        foregroundBrush.Color.A,
                        foregroundBrush.Color.R,
                        foregroundBrush.Color.G,
                        foregroundBrush.Color.B) } }
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

        public static List<ColorItem> CombineClassificationColorList(List<ColorItem> lightColors,
            List<ColorItem> darkColors, List<ColorItem> blueColors)
        {
            var colorList = new List<ColorItem>();

            for (var i = 0; i < lightColors.Count; i++)
            {
                var colorItem = new ColorItem
                {
                    Name = lightColors[i].Name
                };

                colorItem.Colors = lightColors[i].Colors
                    .Union(darkColors[i].Colors)
                    .Union(blueColors[i].Colors)
                    .ToDictionary(pair => pair.Key, pair => pair.Value);

                colorList.Add(colorItem);
            }

            return colorList;
        }
    }
}
