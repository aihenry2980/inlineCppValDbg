using EnvDTE;
using EnvDTE80;
using EnvDTE90a;
using Microsoft.VisualStudio.Shell;
using System;
using System.Threading;

namespace InlineCppVarDbg
{
    internal sealed class DebuggerBridge
    {
        private readonly DTE2 dte2;
        private readonly DebuggerEvents debuggerEvents;
        private readonly _dispDebuggerEvents_Event debuggerEventsEvent;
        private int version;

        public event EventHandler DebugStateChanged;

        public DebuggerBridge(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            dte2 = serviceProvider.GetService(typeof(DTE)) as DTE2;
            if (dte2 == null)
            {
                return;
            }

            debuggerEvents = dte2.Events.DebuggerEvents;
            debuggerEventsEvent = debuggerEvents as _dispDebuggerEvents_Event;
            if (debuggerEventsEvent == null)
            {
                return;
            }

            debuggerEventsEvent.OnEnterBreakMode += OnEnterBreakMode;
            debuggerEventsEvent.OnContextChanged += OnContextChanged;
            debuggerEventsEvent.OnEnterRunMode += OnEnterRunMode;
            debuggerEventsEvent.OnEnterDesignMode += OnEnterDesignMode;
        }

        public int Version => Volatile.Read(ref version);

        public bool IsInDesignMode()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (dte2 == null || dte2.Debugger == null)
            {
                return true;
            }

            try
            {
                return dte2.Debugger.CurrentMode == dbgDebugMode.dbgDesignMode;
            }
            catch
            {
                return true;
            }
        }

        public bool TryGetCurrentBreakContext(out BreakContext context)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            context = default;

            if (dte2 == null || dte2.Debugger == null)
            {
                return false;
            }

            Debugger debugger = dte2.Debugger;
            if (debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
            {
                return false;
            }

            StackFrame2 frame = debugger.CurrentStackFrame as StackFrame2;
            if (frame == null)
            {
                return false;
            }

            string language = frame.Language ?? string.Empty;
            bool isCppLanguage =
                language.IndexOf("C++", StringComparison.OrdinalIgnoreCase) >= 0 ||
                language.IndexOf("C/C++", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!isCppLanguage)
            {
                return false;
            }

            string fileName = frame.FileName;
            int lineNumber = checked((int)frame.LineNumber);
            if (string.IsNullOrWhiteSpace(fileName) || lineNumber <= 0)
            {
                return false;
            }

            context = new BreakContext(debugger, frame, fileName, lineNumber);
            return true;
        }

        public void RaiseExternalInvalidate()
        {
            IncrementVersionAndNotify();
        }

        private void OnEnterBreakMode(dbgEventReason Reason, ref dbgExecutionAction ExecutionAction)
        {
            IncrementVersionAndNotify();
        }

        private void OnContextChanged(Process NewProcess, Program NewProgram, EnvDTE.Thread NewThread, StackFrame NewStackFrame)
        {
            IncrementVersionAndNotify();
        }

        private void OnEnterRunMode(dbgEventReason Reason)
        {
            IncrementVersionAndNotify();
        }

        private void OnEnterDesignMode(dbgEventReason Reason)
        {
            IncrementVersionAndNotify();
        }

        private void IncrementVersionAndNotify()
        {
            Interlocked.Increment(ref version);
            DebugStateChanged?.Invoke(this, EventArgs.Empty);
        }

        internal readonly struct BreakContext
        {
            public BreakContext(Debugger debugger, StackFrame2 stackFrame, string fileName, int lineNumber)
            {
                Debugger = debugger;
                StackFrame = stackFrame;
                FileName = fileName;
                LineNumber = lineNumber;
            }

            public Debugger Debugger { get; }
            public StackFrame2 StackFrame { get; }
            public string FileName { get; }
            public int LineNumber { get; }
        }
    }
}
