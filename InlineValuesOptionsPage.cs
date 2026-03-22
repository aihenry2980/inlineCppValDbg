using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;

namespace InlineCppVarDbg
{
    public sealed class InlineValuesOptionsPage : DialogPage
    {
        private const int DefaultPreviousLineCount = 20;
        private const InlineValueDisplayMode DefaultDisplayMode = InlineValueDisplayMode.Inline;
        private const string DefaultValueBackgroundColor = "#c0f3b9";
        private const string DefaultUninitializedValueBackgroundColor = "#FFF59D";
        private const string DefaultValueChangedAccentColor = "#FF8FB1";
        private const int DefaultValueChipFontSize = 10;
        private int previousLineCount = DefaultPreviousLineCount;
        private InlineValueDisplayMode displayMode = DefaultDisplayMode;
        private Color valueBackgroundColor = ParseHexColor(DefaultValueBackgroundColor, DefaultValueBackgroundColor);
        private Color uninitializedValueBackgroundColor = ParseHexColor(DefaultUninitializedValueBackgroundColor, DefaultUninitializedValueBackgroundColor);
        private Color valueChangedAccentColor = ParseHexColor(DefaultValueChangedAccentColor, DefaultValueChangedAccentColor);
        private int valueChipFontSize = DefaultValueChipFontSize;
        private InlineValueEvaluationKinds evaluationKinds = InlineValueEvaluationKinds.All;

        [Category("General")]
        [DisplayName("Previous line count")]
        [Description("How many lines before the current debug line should show inline values.")]
        public int PreviousLineCount
        {
            get => previousLineCount;
            set => previousLineCount = value;
        }

        [Category("General")]
        [DisplayName("Display mode")]
        [Description("Inline places values after each identifier. End-of-line places value chips at the end of each code line.")]
        public InlineValueDisplayMode DisplayMode
        {
            get => displayMode;
            set => displayMode = value;
        }

        [Category("Appearance")]
        [DisplayName("Value chip background")]
        [Description("Background color for inline value chips.")]
        [Editor(typeof(ColorEditor), typeof(System.Drawing.Design.UITypeEditor))]
        public Color ValueBackgroundColor
        {
            get => valueBackgroundColor;
            set => valueBackgroundColor = value;
        }

        [Category("Appearance")]
        [DisplayName("Uninitialized chip background")]
        [Description("Background color used for uninitialized values (Not Init).")]
        [Editor(typeof(ColorEditor), typeof(System.Drawing.Design.UITypeEditor))]
        public Color UninitializedValueBackgroundColor
        {
            get => uninitializedValueBackgroundColor;
            set => uninitializedValueBackgroundColor = value;
        }

        [Category("Appearance")]
        [DisplayName("Changed value accent")]
        [Description("Accent color used when a value changes after stepping.")]
        [Editor(typeof(ColorEditor), typeof(System.Drawing.Design.UITypeEditor))]
        public Color ValueChangedAccentColor
        {
            get => valueChangedAccentColor;
            set => valueChangedAccentColor = value;
        }

        [Category("Appearance")]
        [DisplayName("Value chip font size")]
        [Description("Font size for inline value chips.")]
        public int ValueChipFontSize
        {
            get => valueChipFontSize;
            set => valueChipFontSize = value;
        }

        [Category("Evaluation")]
        [DisplayName("Basic scalar values")]
        [Description("Evaluate and show basic scalar types such as bool, char, int, float, and double.")]
        public bool EvaluateBasicScalars
        {
            get => HasEvaluationKind(InlineValueEvaluationKinds.BasicScalars);
            set => SetEvaluationKind(InlineValueEvaluationKinds.BasicScalars, value);
        }

        [Category("Evaluation")]
        [DisplayName("Enum values")]
        [Description("Evaluate and show enum variables and enum entries as integers.")]
        public bool EvaluateEnumValues
        {
            get => HasEvaluationKind(InlineValueEvaluationKinds.Enums);
            set => SetEvaluationKind(InlineValueEvaluationKinds.Enums, value);
        }

        [Category("Evaluation")]
        [DisplayName("Array values")]
        [Description("Evaluate and show array values using their front entries only.")]
        public bool EvaluateArrayValues
        {
            get => HasEvaluationKind(InlineValueEvaluationKinds.Arrays);
            set => SetEvaluationKind(InlineValueEvaluationKinds.Arrays, value);
        }

        [Category("Evaluation")]
        [DisplayName("Get/Is function calls")]
        [Description("Evaluate zero-argument Get*() and Is*() calls when their return type is eligible.")]
        public bool EvaluateGetterCalls
        {
            get => HasEvaluationKind(InlineValueEvaluationKinds.GetterCalls);
            set => SetEvaluationKind(InlineValueEvaluationKinds.GetterCalls, value);
        }

        [Category("Evaluation")]
        [DisplayName("Indexed expressions")]
        [Description("Evaluate indexed expressions such as a[b] in addition to showing b itself.")]
        public bool EvaluateIndexedExpressions
        {
            get => HasEvaluationKind(InlineValueEvaluationKinds.IndexedExpressions);
            set => SetEvaluationKind(InlineValueEvaluationKinds.IndexedExpressions, value);
        }

        [Category("Evaluation")]
        [DisplayName("Const/constexpr/macro fallback")]
        [Description("Use fallback debugger expression evaluation for const values, constexprs, and macros when available.")]
        public bool EvaluateFallbackExpressions
        {
            get => HasEvaluationKind(InlineValueEvaluationKinds.FallbackExpressions);
            set => SetEvaluationKind(InlineValueEvaluationKinds.FallbackExpressions, value);
        }

        [Category("Evaluation")]
        [DisplayName("Null pointers")]
        [Description("Show null raw pointers as a null chip. Non-null pointers remain suppressed.")]
        public bool ShowNullPointers
        {
            get => HasEvaluationKind(InlineValueEvaluationKinds.NullPointers);
            set => SetEvaluationKind(InlineValueEvaluationKinds.NullPointers, value);
        }

        [Category("Evaluation")]
        [DisplayName("Uninitialized markers")]
        [Description("Show Not Init chips when the debugger value looks uninitialized.")]
        public bool ShowUninitializedValues
        {
            get => HasEvaluationKind(InlineValueEvaluationKinds.UninitializedValues);
            set => SetEvaluationKind(InlineValueEvaluationKinds.UninitializedValues, value);
        }

        protected override void OnActivate(CancelEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            base.OnActivate(e);

            if (Site is IServiceProvider serviceProvider)
            {
                InlineValuesSettings settings = InlineValuesServiceLocator.GetSettings(serviceProvider);
                previousLineCount = settings.PreviousLineCount;
                displayMode = settings.DisplayMode;
                valueBackgroundColor = ParseHexColor(settings.ValueBackgroundColor, DefaultValueBackgroundColor);
                uninitializedValueBackgroundColor = ParseHexColor(settings.UninitializedValueBackgroundColor, DefaultUninitializedValueBackgroundColor);
                valueChangedAccentColor = ParseHexColor(settings.ValueChangedAccentColor, DefaultValueChangedAccentColor);
                valueChipFontSize = settings.ValueChipFontSize;
                evaluationKinds = settings.EvaluationKinds;
            }
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            base.OnApply(e);

            if (Site is IServiceProvider serviceProvider)
            {
                InlineValuesSettings settings = InlineValuesServiceLocator.GetSettings(serviceProvider);
                settings.SetPreviousLineCount(previousLineCount);
                settings.SetDisplayMode(displayMode);
                settings.SetValueBackgroundColor(ToHex(valueBackgroundColor));
                settings.SetUninitializedValueBackgroundColor(ToHex(uninitializedValueBackgroundColor));
                settings.SetValueChangedAccentColor(ToHex(valueChangedAccentColor));
                settings.SetValueChipFontSize(valueChipFontSize);
                settings.SetEvaluationKinds(evaluationKinds);
            }
        }

        private bool HasEvaluationKind(InlineValueEvaluationKinds kind)
        {
            return (evaluationKinds & kind) == kind;
        }

        private void SetEvaluationKind(InlineValueEvaluationKinds kind, bool enabled)
        {
            if (enabled)
            {
                evaluationKinds |= kind;
            }
            else
            {
                evaluationKinds &= ~kind;
            }
        }

        private static Color ParseHexColor(string value, string fallbackHex)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return ColorTranslator.FromHtml(fallbackHex);
            }

            try
            {
                return ColorTranslator.FromHtml(value.Trim());
            }
            catch
            {
                return ColorTranslator.FromHtml(fallbackHex);
            }
        }

        private static string ToHex(Color color)
        {
            if (color.A >= 255)
            {
                return string.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
            }

            return string.Format("#{0:X2}{1:X2}{2:X2}{3:X2}", color.A, color.R, color.G, color.B);
        }
    }
}
