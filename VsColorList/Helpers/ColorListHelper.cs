using Community.VisualStudio.Toolkit;
using Microsoft.VisualStudio.Language.StandardClassification;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Classification;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using VsColorList.Models;

namespace VsColorList.Helpers
{
    public static class ColorListHelper
    {
        public static List<ColorItem> GetEnvironmentColorList(string themeName)
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

        public static List<ColorItem> GetVsBrushColorList(string themeName)
        {
            var colorList = new List<ColorItem>();

            foreach (var brushProperty in typeof(VsBrushes).GetProperties())
            {
                var name = brushProperty.GetValue(null);
                var brush = Application.Current.Resources[name] as SolidColorBrush;

                if (brush == null)
                {
                    continue;
                }

                var colorItem = new ColorItem
                {
                    Name = brushProperty.Name,
                    Colors = { { themeName, ToDrawingColor(brush) } }
                };

                colorList.Add(colorItem);
            }

            return colorList;
        }

        public static List<KeyValuePair<ThemeResourceKey, uint>> GetVsColorList()
        {
            return VsColors.GetCurrentThemedColorValues().OrderBy(kp => kp.Key.Name).ToList();
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
                    Colors = { { themeName, ToDrawingColor(foregroundBrush) } }
                };

                colorList.Add(colorItem);
            }

            return colorList;
        }

        public static List<ColorItem> CombineEnvironmentColorList(
            List<ColorItem> lightEnvironmentColorList,
            List<ColorItem> darkEnvironmentColorList,
            List<ColorItem> blueEnvironmentColorList)
        {
            var colorList = new List<ColorItem>();

            for (var i = 0; i < lightEnvironmentColorList.Count; i++)
            {
                var colorItem = new ColorItem
                {
                    Key = lightEnvironmentColorList[i].Key,
                    KeyType = lightEnvironmentColorList[i].Key.KeyType,
                    Category = lightEnvironmentColorList[i].Key.Category,
                    Colors = lightEnvironmentColorList[i].Colors
                        .Union(darkEnvironmentColorList[i].Colors)
                        .Union(blueEnvironmentColorList[i].Colors)
                        .ToDictionary(pair => pair.Key, pair => pair.Value)
                };

                colorList.Add(colorItem);
            }

            return colorList;
        }

        public static List<ColorItem> CombineVsColorList(
            List<KeyValuePair<ThemeResourceKey, uint>> lightVsColorList,
            List<KeyValuePair<ThemeResourceKey, uint>> darkVsColorList,
            List<KeyValuePair<ThemeResourceKey, uint>> blueVsColorList)
        {
            var colorList = new List<ColorItem>();

            for (var i = 0; i < lightVsColorList.Count; i++)
            {
                colorList.Add(new ColorItem
                {
                    Key = lightVsColorList[i].Key,
                    KeyType = lightVsColorList[i].Key.KeyType,
                    Category = lightVsColorList[i].Key.Category,
                    Colors = {
                        { "light", VSColorTheme.GetThemedColor(lightVsColorList[i].Key) },
                        { "dark", VSColorTheme.GetThemedColor(darkVsColorList[i].Key) },
                        { "blue", VSColorTheme.GetThemedColor(blueVsColorList[i].Key) },
                    }
                });
            }

            return colorList;
        }

        public static List<ColorItem> CombineClassificationColorList(
            List<ColorItem> lightClassificationList,
            List<ColorItem> darkClassificationList,
            List<ColorItem> blueClassificationList)
        {
            var colorList = new List<ColorItem>();

            for (var i = 0; i < lightClassificationList.Count; i++)
            {
                var colorItem = new ColorItem
                {
                    Name = lightClassificationList[i].Name,
                    Colors = lightClassificationList[i].Colors
                        .Union(darkClassificationList[i].Colors)
                        .Union(blueClassificationList[i].Colors)
                        .ToDictionary(pair => pair.Key, pair => pair.Value)
                };

                colorList.Add(colorItem);
            }

            return colorList;
        }

        private static System.Drawing.Color ToDrawingColor(SolidColorBrush solidColorBrush)
        {
            return System.Drawing.Color.FromArgb(
                solidColorBrush.Color.A,
                solidColorBrush.Color.R,
                solidColorBrush.Color.G,
                solidColorBrush.Color.B);
        }
    }
}
