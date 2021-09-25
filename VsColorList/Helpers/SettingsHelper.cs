using Community.VisualStudio.Toolkit;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

namespace VsColorList.Helpers
{
    public static class SettingsHelper
    {
        public static async Task<string> Export()
        {
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()) + ".vssettings";
            await KnownCommands.Tools_ImportandExportSettings.ExecuteAsync($@"/export:""{path}""");
            return path;
        }

        public static async Task Import(string path)
        {
            var resourcePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", path);
            await KnownCommands.Tools_ImportandExportSettings.ExecuteAsync($@"/import:""{resourcePath}""");
        }
    }
}
