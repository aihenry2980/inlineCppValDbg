using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace InlineCppVarDbg
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideOptionPage(typeof(InlineValuesOptionsPage), "Inline Cpp Var Dbg", "General", 0, 0, true)]
    [Guid(PackageGuidString)]
    public sealed class InlineCppVarDbgPackage : AsyncPackage
    {
        public const string PackageGuidString = "63e068a1-5c78-4229-9118-88306593cf1a";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            InlineValuesServiceLocator.EnsureInitialized(this);
            await ToggleInlineValuesCommand.InitializeAsync(this);
            await ToggleValueDisplayModeCommand.InitializeAsync(this);
        }
    }
}
