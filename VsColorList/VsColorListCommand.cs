using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Community.VisualStudio.Toolkit;
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
            var lightColors = VsColors.GetCurrentThemedColorValues().OrderBy(kp => kp.Key.Name).ToList();
            var envLightList = ColorListHelper.GetColorList("light");

            // Dark
            await SettingsHelper.Import("Dark.vssettings");
            var darkColors = VsColors.GetCurrentThemedColorValues().OrderBy(kp => kp.Key.Name).ToList();
            var envDarkList = ColorListHelper.GetColorList("dark");

            // Blue
            await SettingsHelper.Import("Blue.vssettings");
            var blueColors = VsColors.GetCurrentThemedColorValues().OrderBy(kp => kp.Key.Name).ToList();
            var envBlueList = ColorListHelper.GetColorList("blue");

            await SettingsHelper.Import(tempBackupPath);

            // VS Colors
            var vsColorsList = ColorListHelper.GetVsColorList(lightColors, darkColors, blueColors);

            // Environment Colors
            var environmentColorsList = ColorListHelper.GetEnvironmentColorList(envLightList, envDarkList, envBlueList);

            // Write to Excel
            var filePath = ExcelHelper.WriteToExcel(vsColorsList, environmentColorsList);

            // Open Excel
            Process.Start(filePath);
        }
    }
}
