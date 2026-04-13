using EnvDTE;
using EnvDTE80;
using EnvDTE90a;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace InlineCppVarDbg
{
    internal sealed class DebuggerBridge
    {
        private readonly DTE2 dte2;
        private readonly DebuggerEvents debuggerEvents;
        private readonly _dispDebuggerEvents_Event debuggerEventsEvent;
        private int version;
        private int manualGetterFunctionSweepRequested;
        private int manualVariableFunctionSweepRequested;
        private int profileNextEvaluationRequestKind;
        private int getterDiagnosticsNextEvaluationRequested;
        private int watchGetterFunctionSweepRequested;

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

        public bool IsManualGetterFunctionSweepRequested => Volatile.Read(ref manualGetterFunctionSweepRequested) != 0;

        public bool IsManualVariableFunctionSweepRequested => Volatile.Read(ref manualVariableFunctionSweepRequested) != 0;

        public bool IsWatchGetterFunctionSweepRequested => Volatile.Read(ref watchGetterFunctionSweepRequested) != 0;

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

        public void RequestProfileForNextEvaluation(ProfileRequestKind requestKind)
        {
            Interlocked.Exchange(ref profileNextEvaluationRequestKind, (int)requestKind);
        }

        public void RequestManualGetterFunctionSweep()
        {
            Interlocked.Exchange(ref manualGetterFunctionSweepRequested, 1);
            IncrementVersionAndNotify();
        }

        public void RequestManualGetterAtCaret()
        {
            Interlocked.Exchange(ref manualGetterFunctionSweepRequested, 1);
            IncrementVersionAndNotify();
        }

        public void RequestGetterDiagnosticsForNextEvaluation()
        {
            Interlocked.Exchange(ref getterDiagnosticsNextEvaluationRequested, 1);
        }

        public void RequestWatchGetterFunctionSweep()
        {
            Interlocked.Exchange(ref watchGetterFunctionSweepRequested, 1);
            IncrementVersionAndNotify();
        }

        public void RequestManualVariableFunctionSweep()
        {
            Interlocked.Exchange(ref manualVariableFunctionSweepRequested, 1);
            IncrementVersionAndNotify();
        }

        public bool TryConsumeProfileForNextEvaluation(out ProfileRequestKind requestKind)
        {
            int rawValue = Interlocked.Exchange(ref profileNextEvaluationRequestKind, (int)ProfileRequestKind.None);
            requestKind = rawValue >= (int)ProfileRequestKind.GetButton && rawValue <= (int)ProfileRequestKind.ToggleOnButton
                ? (ProfileRequestKind)rawValue
                : ProfileRequestKind.None;
            return requestKind != ProfileRequestKind.None;
        }

        public bool TryConsumeGetterDiagnosticsForNextEvaluation()
        {
            return Interlocked.Exchange(ref getterDiagnosticsNextEvaluationRequested, 0) != 0;
        }

        public bool TryConsumeManualGetterAtCaretRequest()
        {
            return Interlocked.Exchange(ref manualGetterFunctionSweepRequested, 0) != 0;
        }

        public bool TryConsumeWatchGetterFunctionSweep()
        {
            return Interlocked.Exchange(ref watchGetterFunctionSweepRequested, 0) != 0;
        }

        public WatchAddResult AddSelectedExpressionsToWatch(IEnumerable<string> expressionLabels, Action<string> selectExpression)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var failures = new List<string>();
            if (dte2 == null)
            {
                return new WatchAddResult(0, 0, new[] { "Visual Studio automation object is unavailable." });
            }

            string[] distinctExpressions = expressionLabels?
                .Where(expression => !string.IsNullOrWhiteSpace(expression))
                .Select(expression => expression.Trim())
                .Distinct(StringComparer.Ordinal)
                .ToArray() ?? Array.Empty<string>();

            if (distinctExpressions.Length == 0)
            {
                return new WatchAddResult(0, 0, Array.Empty<string>());
            }

            try
            {
                dte2.ExecuteCommand("Debug.Watch1");
            }
            catch
            {
                // The AddWatch command can still succeed even if the Watch 1 window was already unavailable to open.
            }

            int added = 0;
            foreach (string expression in distinctExpressions)
            {
                try
                {
                    selectExpression?.Invoke(expression);
                    dte2.ExecuteCommand("Debug.AddWatch");
                    added++;
                }
                catch (Exception ex)
                {
                    failures.Add(expression + " -> " + ex.Message);
                }
            }

            return new WatchAddResult(distinctExpressions.Length, added, failures);
        }

        private void OnEnterBreakMode(dbgEventReason Reason, ref dbgExecutionAction ExecutionAction)
        {
            Interlocked.Exchange(ref manualGetterFunctionSweepRequested, 0);
            Interlocked.Exchange(ref manualVariableFunctionSweepRequested, 0);
            Interlocked.Exchange(ref profileNextEvaluationRequestKind, (int)ProfileRequestKind.None);
            Interlocked.Exchange(ref getterDiagnosticsNextEvaluationRequested, 0);
            Interlocked.Exchange(ref watchGetterFunctionSweepRequested, 0);
            IncrementVersionAndNotify();
        }

        private void OnContextChanged(Process NewProcess, Program NewProgram, EnvDTE.Thread NewThread, StackFrame NewStackFrame)
        {
            Interlocked.Exchange(ref manualGetterFunctionSweepRequested, 0);
            Interlocked.Exchange(ref manualVariableFunctionSweepRequested, 0);
            Interlocked.Exchange(ref profileNextEvaluationRequestKind, (int)ProfileRequestKind.None);
            Interlocked.Exchange(ref getterDiagnosticsNextEvaluationRequested, 0);
            Interlocked.Exchange(ref watchGetterFunctionSweepRequested, 0);
            IncrementVersionAndNotify();
        }

        private void OnEnterRunMode(dbgEventReason Reason)
        {
            Interlocked.Exchange(ref manualGetterFunctionSweepRequested, 0);
            Interlocked.Exchange(ref manualVariableFunctionSweepRequested, 0);
            Interlocked.Exchange(ref profileNextEvaluationRequestKind, (int)ProfileRequestKind.None);
            Interlocked.Exchange(ref getterDiagnosticsNextEvaluationRequested, 0);
            Interlocked.Exchange(ref watchGetterFunctionSweepRequested, 0);
            IncrementVersionAndNotify();
        }

        private void OnEnterDesignMode(dbgEventReason Reason)
        {
            Interlocked.Exchange(ref manualGetterFunctionSweepRequested, 0);
            Interlocked.Exchange(ref manualVariableFunctionSweepRequested, 0);
            Interlocked.Exchange(ref profileNextEvaluationRequestKind, (int)ProfileRequestKind.None);
            Interlocked.Exchange(ref getterDiagnosticsNextEvaluationRequested, 0);
            Interlocked.Exchange(ref watchGetterFunctionSweepRequested, 0);
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

        internal enum ProfileRequestKind
        {
            None = 0,
            GetButton = 1,
            ToggleOnButton = 2,
        }

        internal readonly struct WatchAddResult
        {
            public WatchAddResult(int requestedCount, int addedCount, IReadOnlyList<string> failures)
            {
                RequestedCount = requestedCount;
                AddedCount = addedCount;
                Failures = failures ?? Array.Empty<string>();
            }

            public int RequestedCount { get; }
            public int AddedCount { get; }
            public IReadOnlyList<string> Failures { get; }
        }
    }
}
