using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using VsColorList.Models;

namespace VsColorList
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ListCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("47543af7-47bd-4b5c-b045-30d0aa9d3d78");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ListCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ListCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ListCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new ListCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var vsColorsList = new List<ColorListItem>();
            var environmentColorsList = new List<ColorListItem>();
            var vsLightList = new List<ColorListItem>();
            var vsDarkList = new List<ColorListItem>();
            var vsBlueList = new List<ColorListItem>();
            var tempBackupPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()) + ".vssettings";
            var excelFilePath = Path.Combine(Path.GetTempPath(), "VsColorList") + ".xlsx";

            ExportVsSettings(tempBackupPath);

            // Light
            ImportVsSettings("Light.vssettings");
            var lightColors = VsColors.GetCurrentThemedColorValues().OrderBy(kp => kp.Key.Name).ToList();

            foreach (var color in typeof(EnvironmentColors).GetProperties())
            {
                var themeResourceKey = color.GetValue(null) as ThemeResourceKey;

                if (themeResourceKey != null)
                {
                    vsLightList.Add(new ColorListItem
                    {
                        Key = themeResourceKey,
                        LightColor = VSColorTheme.GetThemedColor(themeResourceKey)
                    });
                }
            }

            // Dark
            ImportVsSettings("Dark.vssettings");
            var darkColors = VsColors.GetCurrentThemedColorValues().OrderBy(kp => kp.Key.Name).ToList();

            foreach (var color in typeof(EnvironmentColors).GetProperties())
            {
                var themeResourceKey = color.GetValue(null) as ThemeResourceKey;

                if (themeResourceKey != null)
                {
                    vsDarkList.Add(new ColorListItem
                    {
                        Key = themeResourceKey,
                        DarkColor = VSColorTheme.GetThemedColor(themeResourceKey)
                    });
                }
            }

            // Blue
            ImportVsSettings("Blue.vssettings");
            var blueColors = VsColors.GetCurrentThemedColorValues().OrderBy(kp => kp.Key.Name).ToList();

            foreach (var color in typeof(EnvironmentColors).GetProperties())
            {
                var themeResourceKey = color.GetValue(null) as ThemeResourceKey;

                if (themeResourceKey != null)
                {
                    vsBlueList.Add(new ColorListItem
                    {
                        Key = themeResourceKey,
                        BlueColor = VSColorTheme.GetThemedColor(themeResourceKey)
                    });
                }
            }

            ImportVsSettings(tempBackupPath);

            for (var i = 0; i < lightColors.Count; i++)
            {
                vsColorsList.Add(new ColorListItem
                {
                    Key = lightColors[i].Key,
                    KeyType = lightColors[i].Key.KeyType,
                    Category = lightColors[i].Key.Category,
                    LightColor = VSColorTheme.GetThemedColor(lightColors[i].Key),
                    DarkColor = VSColorTheme.GetThemedColor(darkColors[i].Key),
                    BlueColor = VSColorTheme.GetThemedColor(blueColors[i].Key),
                });
            }

            for (var i = 0; i < vsLightList.Count; i++)
            {
                environmentColorsList.Add(new ColorListItem
                {
                    Key = vsLightList[i].Key,
                    KeyType = vsLightList[i].Key.KeyType,
                    Category = vsLightList[i].Key.Category,
                    LightColor = vsLightList[i].LightColor,
                    DarkColor = vsDarkList[i].DarkColor,
                    BlueColor = vsBlueList[i].BlueColor
                });
            }

            var vsColorsFile = WriteToExcel(vsColorsList, environmentColorsList, excelFilePath);

            Process.Start(excelFilePath);
        }

        private string WriteToExcel(List<ColorListItem> colors, List<ColorListItem> environmentColors, string path)
        {
            var newFile = new FileInfo(path);
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(path);
            }

            using (var package = new ExcelPackage(newFile))
            {
                AddWorkSheet(package, environmentColors, "EnvironmentColors");
                AddWorkSheet(package, colors, "VsColors");             

                package.Save();
            }

            return newFile.FullName;
        }

        private void AddWorkSheet(ExcelPackage package, List<ColorListItem> colors, string title)
        {
            // Add a new worksheet to the empty workbook
            var worksheet = package.Workbook.Worksheets.Add(title);

            //Add the headers
            worksheet.Cells[1, 1].Value = "Key";
            worksheet.Cells[1, 2].Value = "Light";
            worksheet.Cells[1, 3].Value = "Dark";
            worksheet.Cells[1, 4].Value = "Blue";
            worksheet.Cells[1, 5].Value = "Light Hex ARGB";
            worksheet.Cells[1, 6].Value = "Dark Hex ARGB";
            worksheet.Cells[1, 7].Value = "Blue Hex ARGB";
            worksheet.Cells[1, 8].Value = "Light RGB";
            worksheet.Cells[1, 9].Value = "Dark RGB";
            worksheet.Cells[1, 10].Value = "Blue RGB";
            worksheet.Cells[1, 11].Value = "Category";
            worksheet.Cells[1, 12].Value = "Key Type";
            worksheet.Cells["A1:L1"].Style.Font.Bold = true;
            worksheet.Cells["A1:L1"].AutoFilter = true;

            for (var i = 0; i < colors.Count; i++)
            {
                worksheet.Cells[$"A{i + 2}"].Value = colors[i].Key;

                worksheet.Cells[$"B{i + 2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"B{i + 2}"].Style.Fill.BackgroundColor.SetColor(colors[i].LightColor);

                worksheet.Cells[$"C{i + 2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"C{i + 2}"].Style.Fill.BackgroundColor.SetColor(colors[i].DarkColor);

                worksheet.Cells[$"D{i + 2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"D{i + 2}"].Style.Fill.BackgroundColor.SetColor(colors[i].BlueColor);

                worksheet.Cells[$"E{i + 2}"].Value = ToHexARGB(colors[i].LightColor);
                worksheet.Cells[$"F{i + 2}"].Value = ToHexARGB(colors[i].DarkColor);
                worksheet.Cells[$"G{i + 2}"].Value = ToHexARGB(colors[i].BlueColor);

                worksheet.Cells[$"H{i + 2}"].Value = ToRGB(colors[i].LightColor);
                worksheet.Cells[$"I{i + 2}"].Value = ToRGB(colors[i].DarkColor);
                worksheet.Cells[$"J{i + 2}"].Value = ToRGB(colors[i].BlueColor);

                worksheet.Cells[$"K{i + 2}"].Value = ToCategoryName(colors[i].Category);

                worksheet.Cells[$"L{i + 2}"].Value = colors[i].KeyType.ToString();
            }

            worksheet.Cells.AutoFitColumns(0);
        }

        private string ToHexARGB(Color color) => $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";

        private string ToRGB(Color color) => $"{color.R},{color.G},{color.B}";

        private string ToCategoryName(Guid guid)
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

        private void ExportVsSettings(string path)
        {
            var dte = ServiceProvider.GetService(typeof(EnvDTE._DTE)) as EnvDTE.DTE;
            dte.ExecuteCommand("Tools.ImportandExportSettings", $@"/export:""{path}""");
        }

        private void ImportVsSettings(string path)
        {
            var dte = ServiceProvider.GetService(typeof(EnvDTE._DTE)) as EnvDTE.DTE;
            var resourcePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", path);
            dte.ExecuteCommand("Tools.ImportandExportSettings", $@"/import:""{resourcePath}""");
        }
    }
}
