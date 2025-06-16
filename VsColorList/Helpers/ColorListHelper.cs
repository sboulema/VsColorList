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
                    Name = color.Name,
                    Key = themeResourceKey,
                    Colors = { { themeName, VSColorTheme.GetThemedColor(themeResourceKey) } },
                    Type = "EnvironmentColor",
                    KeyType = themeResourceKey.KeyType,
                    Category = themeResourceKey.Category,
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

                if (!(Application.Current.Resources[name] is SolidColorBrush brush))
                {
                    continue;
                }

                var colorItem = new ColorItem
                {
                    Name = brushProperty.Name,
                    Colors = { { themeName, ToDrawingColor(brush) } },
                    Type = "VsBrush",
                };

                colorList.Add(colorItem);
            }

            return colorList;
        }

        public static List<ColorItem> GetVsColorList(string themeName)
            => VsColors
                .GetCurrentThemedColorValues()
                .OrderBy(kp => kp.Key.Name)
                .Select(kp => new ColorItem
                {
                    Name = kp.Key.Name,
                    Category = kp.Key.Category,
                    Key = kp.Key,
                    KeyType = kp.Key.KeyType,
                    Type = "VsColor",
                    Colors = { { themeName, System.Drawing.Color.FromArgb((int)kp.Value) } }
                })
                .ToList();

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
                    Colors = { { themeName, ToDrawingColor(foregroundBrush) } },
                    Type = "ClassificationColor",
                };

                colorList.Add(colorItem);
            }

            return colorList;
        }

        public static List<ColorItem> CombineColorLists(
            List<ColorItem> lightColorList,
            List<ColorItem> darkColorList,
            List<ColorItem> blueColorList)
        {
            var colorList = new List<ColorItem>();

            for (var i = 0; i < lightColorList.Count; i++)
            {
                var colorItem = new ColorItem
                {
                    Name = lightColorList[i].Name,
                    Key = lightColorList[i].Key,
                    KeyType = lightColorList[i].Key?.KeyType,
                    Category = lightColorList[i].Key?.Category,
                    Type = lightColorList[i].Type,
                    Colors = lightColorList[i].Colors
                        .Union(darkColorList[i].Colors)
                        .Union(blueColorList[i].Colors)
                        .ToDictionary(pair => pair.Key, pair => pair.Value)
                };

                colorList.Add(colorItem);
            }

            return colorList;
        }

        private static System.Drawing.Color ToDrawingColor(SolidColorBrush solidColorBrush)
            => System.Drawing.Color.FromArgb(
                solidColorBrush.Color.A,
                solidColorBrush.Color.R,
                solidColorBrush.Color.G,
                solidColorBrush.Color.B);
    }
}
