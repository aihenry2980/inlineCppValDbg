using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;

namespace InlineCppVarDbg
{
    internal sealed class ToggleInlineValuesCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("e7737c43-b5a3-4a66-b0a8-b24f4d560d87");

        private readonly AsyncPackage package;

        private ToggleInlineValuesCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package;
            AddCommand(commandService, CommandId);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                new ToggleInlineValuesCommand(package, commandService);
            }
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            InlineValuesSettings settings = InlineValuesServiceLocator.GetSettings(package);
            settings.SetEnabled(!settings.IsEnabled);

            DebuggerBridge bridge = InlineValuesServiceLocator.GetBridge(package);
            bridge.RaiseExternalInvalidate();
        }

        private void AddCommand(OleMenuCommandService commandService, int commandId)
        {
            var menuCommandId = new CommandID(CommandSet, commandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandId);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(sender is OleMenuCommand menuCommand))
            {
                return;
            }

            InlineValuesSettings settings = InlineValuesServiceLocator.GetSettings(package);
            bool enabled = settings.IsEnabled;

            menuCommand.Visible = true;
            menuCommand.Enabled = true;
            menuCommand.Checked = false;
            menuCommand.Text = enabled ? "ON/off" : "on/OFF";
        }
    }
}
