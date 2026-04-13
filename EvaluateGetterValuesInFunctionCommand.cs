using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using System.Windows.Input;

namespace InlineCppVarDbg
{
    internal sealed class EvaluateGetterValuesInFunctionCommand
    {
        public const int CommandId = 0x0103;
        public static readonly Guid CommandSet = new Guid("e7737c43-b5a3-4a66-b0a8-b24f4d560d87");

        private readonly AsyncPackage package;

        private EvaluateGetterValuesInFunctionCommand(AsyncPackage package, OleMenuCommandService commandService)
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
                new EvaluateGetterValuesInFunctionCommand(package, commandService);
            }
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            InlineValuesSettings settings = InlineValuesServiceLocator.GetSettings(package);
            DebuggerBridge bridge = InlineValuesServiceLocator.GetBridge(package);
            bool addToWatch = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            if (addToWatch)
            {
                bridge.RequestWatchGetterFunctionSweep();
                return;
            }

            if (settings.IsEnabled)
            {
                bridge.RequestProfileForNextEvaluation(DebuggerBridge.ProfileRequestKind.GetButton);
                bridge.RequestGetterDiagnosticsForNextEvaluation();
            }

            bridge.RequestManualGetterAtCaret();
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

            DebuggerBridge bridge = InlineValuesServiceLocator.GetBridge(package);
            bool hasBreakContext = bridge.TryGetCurrentBreakContext(out _);

            menuCommand.Visible = true;
            menuCommand.Enabled = hasBreakContext;
            menuCommand.Checked = false;
            menuCommand.Text = "GET";
        }
    }
}
