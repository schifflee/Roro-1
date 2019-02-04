

using System;
using OpenQA.Selenium.IE;
using Microsoft.Win32;
using System.Drawing;

namespace Roro
{
    public sealed class IEContext : WebContext
    {
        public IEContext()
        {
            // session
            var session = Guid.NewGuid().ToString();

            // registry
            this.SetupRegistry();

            // service
            var service = InternetExplorerDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = false;

            // options
            var options = new InternetExplorerOptions();
            options.InitialBrowserUrl = string.Format("{0}:{1}", WebContext.DefaultUrl, session);

            // driver
            this.Driver = new InternetExplorerDriver(service, options, this.Timeout);

            // process
            this.ProcessId = this.GetProcessIdFromSession(session);
        }

        protected override bool UpdateViewport(WinElement winElement)
        {
            this.Viewport = Rectangle.Empty;
            if (winElement != null && winElement.MainWindow is WinElement target && target.ProcessId == this.ProcessId)
            {
                if (target.GetElement(x => x.Class == "Internet Explorer_Server" || x.Class == "NewTabWnd") is WinElement viewport)
                {
                    this.Viewport = viewport.Bounds;
                    return true;
                }
            }
            return false;
        }

        private void SetupRegistry()
        {
            //// ERROR: Unable to get browser
            //// REPLICATE: Start driver, navigate manually to other site
            Registry.SetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BFCACHE", "iexplore.exe", 0, RegistryValueKind.DWord);

            //// ERROR: WebDriverException: Unexpected error launching Internet Explorer. Protected Mode settings are not the same for all zones. Enable Protected Mode must be set to the same value (enabled or disabled) for all zones.
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\0", "2500", 0, RegistryValueKind.DWord);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2500", 0, RegistryValueKind.DWord);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2500", 0, RegistryValueKind.DWord);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2500", 0, RegistryValueKind.DWord);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4", "2500", 0, RegistryValueKind.DWord);
        }
    }
}
