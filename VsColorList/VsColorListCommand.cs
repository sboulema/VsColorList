using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Community.VisualStudio.Toolkit;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json;
using VsColorList.Helpers;

namespace VsColorList
{
    [Command(0x0100)]
    internal sealed class VsColorListCommand : BaseCommand<VsColorListCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var tempBackupPath = await SettingsHelper.Export();

            // Light
            await VS.StatusBar.ShowProgressAsync("Getting Light theme colors", 1, 5);
            await SettingsHelper.Import("Light.vssettings");
            var lightVsBrushList = ColorListHelper.GetVsBrushColorList("light");
            var lightVsColorList = ColorListHelper.GetVsColorList("light");
            var lightEnvironmentColorList = ColorListHelper.GetEnvironmentColorList("light");
            var lightClassificationList = await ColorListHelper.GetClassificationColorList("light");

            // Dark
            await VS.StatusBar.ShowProgressAsync("Getting Dark theme colors", 2, 5);
            await SettingsHelper.Import("Dark.vssettings");
            var darkVsBrushList = ColorListHelper.GetVsBrushColorList("dark");
            var darkVsColorList = ColorListHelper.GetVsColorList("dark");
            var darkEnvironmentColorList = ColorListHelper.GetEnvironmentColorList("dark");
            var darkClassificationList = await ColorListHelper.GetClassificationColorList("dark");

            // Blue
            await VS.StatusBar.ShowProgressAsync("Getting Blue theme colors", 3, 5);
            await SettingsHelper.Import("Blue.vssettings");
            var blueVsBrushList = ColorListHelper.GetVsBrushColorList("blue");
            var blueVsColorList = ColorListHelper.GetVsColorList("blue");
            var blueEnvironmentColorList = ColorListHelper.GetEnvironmentColorList("blue");
            var blueClassificationList = await ColorListHelper.GetClassificationColorList("blue");

            await SettingsHelper.Import(tempBackupPath);

            // VS Brushes
            var vsBrushesList = ColorListHelper
                .CombineColorLists(lightVsBrushList, darkVsBrushList, blueVsBrushList);

            // VS Colors
            var vsColorsList = ColorListHelper
                .CombineColorLists(lightVsColorList, darkVsColorList, blueVsColorList);

            // Environment Colors
            var environmentColorsList = ColorListHelper
                .CombineColorLists(lightEnvironmentColorList, darkEnvironmentColorList, blueEnvironmentColorList);

            // Classification Colors
            var classificationColorsList = ColorListHelper
                .CombineColorLists(lightClassificationList, darkClassificationList, blueClassificationList);

            // Write to Excel
            await VS.StatusBar.ShowProgressAsync("Writing colors to Excel file", 4, 5);
            var filePath = ExcelHelper
                .WriteToExcel(vsBrushesList, vsColorsList, environmentColorsList, classificationColorsList);

            // Write to JSON
            await VS.StatusBar.ShowProgressAsync("Writing colors to JSON file", 5, 5);
            var json = JsonConvert
                .SerializeObject(
                    vsBrushesList
                        .Concat(vsColorsList)
                        .Concat(environmentColorsList)
                        .Concat(classificationColorsList));
            File.WriteAllText(Path.ChangeExtension(filePath, "json"), json);

            await VS.MessageBox.ShowAsync(filePath);
        }
    }
}
