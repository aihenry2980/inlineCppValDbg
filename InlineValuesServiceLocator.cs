using Microsoft.VisualStudio.Shell;
using System;

namespace InlineCppVarDbg
{
    internal static class InlineValuesServiceLocator
    {
        private static readonly object Gate = new object();
        private static DebuggerBridge debuggerBridge;
        private static InlineValuesSettings settings;

        public static void EnsureInitialized(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            lock (Gate)
            {
                if (settings == null)
                {
                    settings = new InlineValuesSettings(serviceProvider);
                }

                if (debuggerBridge == null)
                {
                    debuggerBridge = new DebuggerBridge(serviceProvider);
                }
            }
        }

        public static DebuggerBridge GetBridge(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            EnsureInitialized(serviceProvider);
            return debuggerBridge;
        }

        public static InlineValuesSettings GetSettings(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            EnsureInitialized(serviceProvider);
            return settings;
        }
    }
}
