

using System;
using OpenQA.Selenium.Edge;
using Microsoft.Win32;
using System.Linq;
using System.Drawing;

namespace Roro
{
    public sealed class EdgeContext : WebContext
    {
        public EdgeContext()
        {
            // session
            var session = "-ServerName:MicrosoftEdge";

            // service
            var service = EdgeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = false;

            // options
            var options = new EdgeOptions();

            // driver
            this.Driver = new EdgeDriver(service, options, this.Timeout);

            // process
            this.ProcessId = this.GetProcessIdFromSession(session);
        }

        protected override bool UpdateViewport(WinElement winElement)
        {
            this.Viewport = Rectangle.Empty;
            if (winElement != null && winElement.MainWindow is WinElement mainWindow
                && mainWindow.Class == "ApplicationFrameWindow"
                && mainWindow.Children.FirstOrDefault(x => x.Type == "window" && x.Name == "Microsoft Edge" && x.Class == "Windows.UI.Core.CoreWindow") is WinElement target
                && target.ProcessId == this.ProcessId)
            {
                if (target.GetElement(x => x.Class == "Internet Explorer_Server" || x.Class == "NewTabPage") is WinElement viewport)
                {
                    this.Viewport = viewport.Bounds;
                    return true;
                }
            }
            return false;
        }
    }
}
