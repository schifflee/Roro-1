

using System;
using OpenQA.Selenium.Chrome;
using System.Drawing;

namespace Roro
{
    public sealed class ChromeContext : WebContext
    {
        public ChromeContext()
        {
            // session
            var session = Guid.NewGuid().ToString();
            
            // service
            var service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = false;

            // options
            var options = new ChromeOptions();
            options.AddArgument("--force-renderer-accessibility");
            options.AddArgument(session.ToString());

            // driver
            this.Driver = new ChromeDriver(service, options, this.Timeout);

            // process
            this.ProcessId = this.GetProcessIdFromSession(session);
        }

        protected override bool UpdateViewport(WinElement winElement)
        {
            this.Viewport = Rectangle.Empty;
            if (winElement != null && winElement.MainWindow is WinElement target && target.ProcessId == this.ProcessId)
            {
                if (target.GetElement(x => x.Class == "Chrome_RenderWidgetHostHWND") is WinElement viewport)
                {
                    this.Viewport = viewport.Bounds;
                    return true;
                }
            }
            return false;
        }
    }
}
