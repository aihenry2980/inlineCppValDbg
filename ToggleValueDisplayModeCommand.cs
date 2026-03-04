using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;

namespace InlineCppVarDbg
{
    internal sealed class ToggleValueDisplayModeCommand
    {
        public const int CommandId = 0x0101;
        public static readonly Guid CommandSet = new Guid("e7737c43-b5a3-4a66-b0a8-b24f4d560d87");

        private readonly AsyncPackage package;

        private ToggleValueDisplayModeCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package;

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandId);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                new ToggleValueDisplayModeCommand(package, commandService);
            }
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            InlineValuesSettings settings = InlineValuesServiceLocator.GetSettings(package);
            InlineValueDisplayMode nextMode = settings.DisplayMode == InlineValueDisplayMode.EndOfLine
                ? InlineValueDisplayMode.Inline
                : InlineValueDisplayMode.EndOfLine;
            settings.SetDisplayMode(nextMode);

            DebuggerBridge bridge = InlineValuesServiceLocator.GetBridge(package);
            bridge.RaiseExternalInvalidate();
        }

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var menuCommand = sender as OleMenuCommand;
            if (menuCommand == null)
            {
                return;
            }

            InlineValuesSettings settings = InlineValuesServiceLocator.GetSettings(package);
            bool isEndOfLineMode = settings.DisplayMode == InlineValueDisplayMode.EndOfLine;
            menuCommand.Checked = isEndOfLineMode;
            menuCommand.Text = "Toggle inline/end of line";
            menuCommand.Enabled = true;
            menuCommand.Visible = true;
        }
    }
}
