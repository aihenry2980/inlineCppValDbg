using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;
using System;
using System.ComponentModel.Composition;

namespace InlineCppVarDbg
{
    [Export(typeof(IViewTaggerProvider))]
    [ContentType("text")]
    [TagType(typeof(IntraTextAdornmentTag))]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    [TextViewRole(PredefinedTextViewRoles.Structured)]
    internal sealed class InlineValueTaggerProvider : IViewTaggerProvider
    {
        [Import(typeof(SVsServiceProvider))]
        internal IServiceProvider ServiceProvider { get; set; }

        [Import]
        internal ITextDocumentFactoryService TextDocumentFactoryService { get; set; }

        [Import]
        internal IClassifierAggregatorService ClassifierAggregatorService { get; set; }

        public ITagger<T> CreateTagger<T>(ITextView textView, ITextBuffer buffer) where T : ITag
        {
            if (!(textView is IWpfTextView wpfTextView))
            {
                return null;
            }

            if (buffer != textView.TextBuffer)
            {
                return null;
            }

            if (!TextDocumentFactoryService.TryGetTextDocument(buffer, out ITextDocument textDocument))
            {
                return null;
            }

            DebuggerBridge bridge = null;
            InlineValuesSettings settings = null;
            ThreadHelper.JoinableTaskFactory.Run(
                async delegate
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    InlineValuesServiceLocator.EnsureInitialized(ServiceProvider);
                    bridge = InlineValuesServiceLocator.GetBridge(ServiceProvider);
                    settings = InlineValuesServiceLocator.GetSettings(ServiceProvider);
                });

            IClassifier classifier = ClassifierAggregatorService?.GetClassifier(buffer);
            return new InlineValueTagger(wpfTextView, buffer, textDocument.FilePath, bridge, settings, classifier) as ITagger<T>;
        }
    }
}
