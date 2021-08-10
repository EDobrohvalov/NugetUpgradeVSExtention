using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using System;
using System.Threading.Tasks;

namespace NugetUpgrade
{
    internal class Logger
    {
        private string _name;
        private IVsOutputWindowPane _pane;
        private readonly IVsOutputWindow output;
        public static Logger Instance;

        public static async System.Threading.Tasks.Task InitializeAsync(AsyncPackage package)
        {
            var output = package.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;
            Instance = new Logger(output);
        }

        public Logger(IVsOutputWindow output)
        {
            this.output = output;
        }

        public async System.Threading.Tasks.Task LogAsync(object message)
        {
            try
            {
                if (await EnsurePaneAsync())
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    _pane.OutputStringThreadSafe(DateTime.Now + ": " + message + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write(ex);
            }
        }

        private async Task<bool> EnsurePaneAsync()
        {
            if (_pane == null)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                if (_pane == null)
                {
                    Guid guid = Guid.NewGuid();
                   
                    output.CreatePane(ref guid, _name, 1, 1);
                    output.GetPane(ref guid, out _pane);
                }

            }

            return _pane != null;
        }
    }
}
