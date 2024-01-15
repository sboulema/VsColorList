using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Media;
using Community.VisualStudio.Toolkit;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.Shell;
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
            await SettingsHelper.Import("Light.vssettings");
            var lightVsBrushList = ColorListHelper.GetVsBrushColorList("light");
            var lightVsColorList = ColorListHelper.GetVsColorList();
            var lightEnvironmentColorList = ColorListHelper.GetEnvironmentColorList("light");
            var lightClassificationList = await ColorListHelper.GetClassificationColorList("light");

            // Dark
            await SettingsHelper.Import("Dark.vssettings");
            var darkVsBrushList = ColorListHelper.GetVsBrushColorList("dark");
            var darkVsColorList = ColorListHelper.GetVsColorList();
            var darkEnvironmentColorList = ColorListHelper.GetEnvironmentColorList("dark");
            var darkClassificationList = await ColorListHelper.GetClassificationColorList("dark");

            // Blue
            await SettingsHelper.Import("Blue.vssettings");
            var blueVsBrushList = ColorListHelper.GetVsBrushColorList("blue");
            var blueVsColorList = ColorListHelper.GetVsColorList();
            var blueEnvironmentColorList = ColorListHelper.GetEnvironmentColorList("blue");
            var blueClassificationList = await ColorListHelper.GetClassificationColorList("blue");

            await SettingsHelper.Import(tempBackupPath);

            // VS Brushes
            var vsBrushesList = ColorListHelper
                .CombineClassificationColorList(lightVsBrushList, darkVsBrushList, blueVsBrushList);

            // VS Colors
            var vsColorsList = ColorListHelper
                .CombineVsColorList(lightVsColorList, darkVsColorList, blueVsColorList);

            // Environment Colors
            var environmentColorsList = ColorListHelper
                .CombineEnvironmentColorList(lightEnvironmentColorList, darkEnvironmentColorList, blueEnvironmentColorList);

            // Classification Colors
            var classificationColorsList = ColorListHelper
                .CombineClassificationColorList(lightClassificationList, darkClassificationList, blueClassificationList);

            // Write to Excel
            var filePath = ExcelHelper
                .WriteToExcel(vsBrushesList, vsColorsList, environmentColorsList, classificationColorsList);

            await VS.MessageBox.ShowAsync(filePath);
        }
    }
}
