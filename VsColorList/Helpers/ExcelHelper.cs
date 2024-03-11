using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using VsColorList.Models;

namespace VsColorList.Helpers
{
    public static class ExcelHelper
    {
        public static string WriteToExcel(
            List<ColorItem> vsBrushesList,
            List<ColorItem> vsColors,
            List<ColorItem> environmentColors,
            List<ColorItem> classificationColors)
        {
            var path = Path.Combine(Path.GetTempPath(), "VsColorList") + ".xlsx";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (var workbook = new XLWorkbook())
            {
                AddWorkSheet(workbook, vsBrushesList, "VsBrushes");
                AddWorkSheet(workbook, environmentColors, "EnvironmentColors");
                AddWorkSheet(workbook, vsColors, "VsColors");
                AddWorkSheet(workbook, classificationColors, "ClassificationColors");

                workbook.SaveAs(path);
            }

            return path;
        }

        private static void AddWorkSheet(IXLWorkbook workbook, List<ColorItem> colors, string title)
        {
            // Add a new worksheet to the empty workbook
            var worksheet = workbook.Worksheets.Add(title);

            //Add the headers
            AddHeaderRow(worksheet);

            for (var i = 0; i < colors.Count; i++)
            {
                AddRow(worksheet, colors[i], i + 2);
            }

            worksheet.Columns().AdjustToContents();
        }

        private static void AddHeaderRow(IXLWorksheet worksheet)
        {
            worksheet.Cell(1, 1).Value = "Key";
            worksheet.Cell(1, 2).Value = "Light";
            worksheet.Cell(1, 3).Value = "Dark";
            worksheet.Cell(1, 4).Value = "Blue";
            worksheet.Cell(1, 5).Value = "Light Hex ARGB";
            worksheet.Cell(1, 6).Value = "Dark Hex ARGB";
            worksheet.Cell(1, 7).Value = "Blue Hex ARGB";
            worksheet.Cell(1, 8).Value = "Light RGB";
            worksheet.Cell(1, 9).Value = "Dark RGB";
            worksheet.Cell(1, 10).Value = "Blue RGB";
            worksheet.Cell(1, 11).Value = "Category";
            worksheet.Cell(1, 12).Value = "Key Type";

            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).SetAutoFilter();
        }

        private static void AddRow(IXLWorksheet worksheet, ColorItem colorItem, int row)
        {
            worksheet.Cell(row, 1).Value = GetIdentifier(colorItem);

            worksheet.Cell(row, 2).Style.Fill.PatternType = XLFillPatternValues.Solid;
            worksheet.Cell(row, 2).Style.Fill.BackgroundColor = ToXLColor(colorItem.Colors["light"]);

            worksheet.Cell(row, 3).Style.Fill.PatternType = XLFillPatternValues.Solid;
            worksheet.Cell(row, 3).Style.Fill.BackgroundColor = ToXLColor(colorItem.Colors["dark"]);

            worksheet.Cell(row, 4).Style.Fill.PatternType = XLFillPatternValues.Solid;
            worksheet.Cell(row, 4).Style.Fill.BackgroundColor = ToXLColor(colorItem.Colors["blue"]);

            worksheet.Cell(row, 5).Value = ToHexARGB(colorItem.Colors["light"]);
            worksheet.Cell(row, 6).Value = ToHexARGB(colorItem.Colors["dark"]);
            worksheet.Cell(row, 7).Value = ToHexARGB(colorItem.Colors["blue"]);

            worksheet.Cell(row, 8).Value = ToRGB(colorItem.Colors["light"]);
            worksheet.Cell(row, 9).Value = ToRGB(colorItem.Colors["dark"]);
            worksheet.Cell(row, 10).Value = ToRGB(colorItem.Colors["blue"]);

            worksheet.Cell(row, 11).Value = ToCategoryName(colorItem.Category);

            worksheet.Cell(row, 12).Value = colorItem.KeyType.ToString();
        }

        private static string ToHexARGB(Color color) => $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";

        private static string ToRGB(Color color) => $"{color.R},{color.G},{color.B}";

        private static XLColor ToXLColor(Color color) => XLColor.FromArgb(color.ToArgb());

        private static string ToCategoryName(Guid? guid)
        {
            switch (guid.ToString())
            {
                case "5d42b198-efca-431c-92aa-8b595d0d39c2":
                    return "User Notifications";
                case "624ed9c3-bdfd-41fa-96c3-7c824ea32e3d":
                    return "Environment";
                case "0cd5aa2b-ef23-4997-80b5-7d0e8fe5b312":
                    return "Graph Document Colors";
                case "92d153ee-57d7-431f-a739-0931ca3f7f70":
                    return "Cider";
                case "92ecf08e-8b13-4cf4-99e9-ae2692382185":
                    return "Tree View";
                case "f1095fad-881f-45f1-8580-589e10325eb8":
                    return "Search Control";
                case "4aff231b-f28a-44f0-a66b-1beeb17cb920":
                    return "Team Explorer";
                case "2138d120-456d-425e-80b5-88d2401fca23":
                    return "Work Item Editor";
                case "b239f458-9f75-4376-959b-4d48b89337f4":
                    return "Manifest Designer";
                default:
                    return guid.ToString();
            }
        }

        private static string GetIdentifier(ColorItem colorItem)
        {
            if (colorItem.Key == null)
            {
                return colorItem.Name;
            }

            return colorItem.Key.ToString();
        }
    }
}
