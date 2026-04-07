using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.ComponentModel.Design;
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
        private const InlineValueTypeRuleKinds DefaultTypeRuleKinds =
            InlineValueTypeRuleKinds.BooleanScalars |
            InlineValueTypeRuleKinds.CharacterScalars |
            InlineValueTypeRuleKinds.IntegralScalars |
            InlineValueTypeRuleKinds.FloatingPointScalars |
            InlineValueTypeRuleKinds.EnumValues |
            InlineValueTypeRuleKinds.BooleanArrays |
            InlineValueTypeRuleKinds.CharacterArrays |
            InlineValueTypeRuleKinds.IntegralArrays |
            InlineValueTypeRuleKinds.FloatingPointArrays |
            InlineValueTypeRuleKinds.EnumArrays |
            InlineValueTypeRuleKinds.SignedCharArrays |
            InlineValueTypeRuleKinds.UnsignedCharArrays |
            InlineValueTypeRuleKinds.SignedIntegralArrays |
            InlineValueTypeRuleKinds.UnsignedIntegralArrays |
            InlineValueTypeRuleKinds.IntegralPointers |
            InlineValueTypeRuleKinds.Unsigned8BitPointers |
            InlineValueTypeRuleKinds.Unsigned16BitPointers |
            InlineValueTypeRuleKinds.Unsigned32BitPointers;
        private int previousLineCount = DefaultPreviousLineCount;
        private InlineValueDisplayMode displayMode = DefaultDisplayMode;
        private Color valueBackgroundColor = ParseHexColor(DefaultValueBackgroundColor, DefaultValueBackgroundColor);
        private Color uninitializedValueBackgroundColor = ParseHexColor(DefaultUninitializedValueBackgroundColor, DefaultUninitializedValueBackgroundColor);
        private Color valueChangedAccentColor = ParseHexColor(DefaultValueChangedAccentColor, DefaultValueChangedAccentColor);
        private int valueChipFontSize = DefaultValueChipFontSize;
        private InlineValueEvaluationKinds evaluationKinds =
            InlineValueEvaluationKinds.BasicScalars |
            InlineValueEvaluationKinds.Enums |
            InlineValueEvaluationKinds.Arrays |
            InlineValueEvaluationKinds.IndexedExpressions |
            InlineValueEvaluationKinds.FallbackExpressions |
            InlineValueEvaluationKinds.NullPointers |
            InlineValueEvaluationKinds.UninitializedValues;
        private InlineValueRuleKinds ruleKinds =
            InlineValueRuleKinds.IntegralPointers |
            InlineValueRuleKinds.ParameterAnchors;
        private InlineValueTypeRuleKinds typeRuleKinds = DefaultTypeRuleKinds;
        private string customRulesText = string.Empty;

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
        [Description("Automatically evaluate zero-argument Get*() and Is*() calls. When disabled, Ctrl+Click a getter in the editor to evaluate it manually for the current break state.")]
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

        [Category("Rules")]
        [DisplayName("Integral pointer master switch")]
        [Description("Master enable for integral pointer dereference formatting. The detailed Integral pointers checkbox must also be enabled.")]
        public bool ShowIntegralPointers
        {
            get => HasRuleKind(InlineValueRuleKinds.IntegralPointers);
            set => SetRuleKind(InlineValueRuleKinds.IntegralPointers, value);
        }

        [Category("Rules")]
        [DisplayName("Parameter anchors")]
        [Description("Show parameter values near the function signature even when the lookback window does not reach that line.")]
        public bool ShowParameterAnchors
        {
            get => HasRuleKind(InlineValueRuleKinds.ParameterAnchors);
            set => SetRuleKind(InlineValueRuleKinds.ParameterAnchors, value);
        }

        [Category("Rules")]
        [DisplayName("Inactive preprocessor code")]
        [Description("Parse values in inactive #if/#ifdef regions. Off is recommended.")]
        public bool ParseInactivePreprocessorCode
        {
            get => HasRuleKind(InlineValueRuleKinds.ParseInactivePreprocessor);
            set => SetRuleKind(InlineValueRuleKinds.ParseInactivePreprocessor, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Bool scalars")]
        [Description("Show plain bool values. Applies only when Basic scalar values is enabled.")]
        public bool ShowBooleanScalars
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.BooleanScalars);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.BooleanScalars, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Char scalars")]
        [Description("Show char, wchar_t, char8_t, char16_t, and char32_t values. Applies only when Basic scalar values is enabled.")]
        public bool ShowCharacterScalars
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.CharacterScalars);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.CharacterScalars, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Integral scalars")]
        [Description("Show integral scalar values such as short, int, long, size_t, and ptrdiff_t. Applies only when Basic scalar values is enabled.")]
        public bool ShowIntegralScalars
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.IntegralScalars);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.IntegralScalars, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Floating scalars")]
        [Description("Show float and double values. Applies only when Basic scalar values is enabled.")]
        public bool ShowFloatingPointScalars
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.FloatingPointScalars);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.FloatingPointScalars, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Enum values")]
        [Description("Show enum values. Applies only when Enum values is enabled.")]
        public bool ShowDetailedEnumValues
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.EnumValues);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.EnumValues, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Bool arrays")]
        [Description("Show bool arrays using a short inline preview. Applies only when Array values is enabled.")]
        public bool ShowBooleanArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.BooleanArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.BooleanArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Char/string arrays")]
        [Description("Show char-like arrays using a short character or string preview. Applies only when Array values is enabled.")]
        public bool ShowCharacterArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.CharacterArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.CharacterArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Signed char arrays")]
        [Description("Show signed char[] and int8_t[] arrays. Applies only when Array values and Char/string arrays are enabled.")]
        public bool ShowSignedCharArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.SignedCharArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.SignedCharArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Unsigned char arrays")]
        [Description("Show unsigned char[] and uint8_t[] arrays. Applies only when Array values and Char/string arrays are enabled.")]
        public bool ShowUnsignedCharArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.UnsignedCharArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.UnsignedCharArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Integral arrays")]
        [Description("Show arrays of integral element types such as int[] or size_t[]. Applies only when Array values is enabled.")]
        public bool ShowIntegralArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.IntegralArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.IntegralArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Signed integral arrays")]
        [Description("Show signed integral arrays such as int[], short[], and long[]. Applies only when Array values and Integral arrays are enabled.")]
        public bool ShowSignedIntegralArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.SignedIntegralArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.SignedIntegralArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Unsigned integral arrays")]
        [Description("Show unsigned integral arrays such as unsigned int[], uint32_t[], and size_t[]. Applies only when Array values and Integral arrays are enabled.")]
        public bool ShowUnsignedIntegralArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.UnsignedIntegralArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.UnsignedIntegralArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Floating arrays")]
        [Description("Show arrays of float or double elements. Applies only when Array values is enabled.")]
        public bool ShowFloatingPointArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.FloatingPointArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.FloatingPointArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Enum arrays")]
        [Description("Show arrays of enum elements. Applies only when Array values is enabled.")]
        public bool ShowEnumArrays
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.EnumArrays);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.EnumArrays, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Bool pointers")]
        [Description("Show bool* values as address plus dereferenced bool. Applies only when Basic scalar values is enabled.")]
        public bool ShowBooleanPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.BooleanPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.BooleanPointers, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Char/string pointers")]
        [Description("Show char-like pointers using a short string preview when possible.")]
        public bool ShowCharacterPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.CharacterPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.CharacterPointers, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Integral pointers")]
        [Description("Show integer-like pointers as address plus dereferenced value, for example int* or uint32_t*. Applies only when Basic scalar values is enabled.")]
        public bool ShowDetailedIntegralPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.IntegralPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.IntegralPointers, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("uint8* pointers")]
        [Description("Show uint8* and uint8_t* values as address plus dereferenced byte value. Applies only when Basic scalar values and Integral pointer master switch are enabled.")]
        public bool ShowUnsigned8BitPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.Unsigned8BitPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.Unsigned8BitPointers, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("uint16* pointers")]
        [Description("Show uint16* and uint16_t* values as address plus dereferenced 16-bit value. Applies only when Basic scalar values and Integral pointer master switch are enabled.")]
        public bool ShowUnsigned16BitPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.Unsigned16BitPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.Unsigned16BitPointers, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("uint32* pointers")]
        [Description("Show uint32* and uint32_t* values as address plus dereferenced 32-bit value. Applies only when Basic scalar values and Integral pointer master switch are enabled.")]
        public bool ShowUnsigned32BitPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.Unsigned32BitPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.Unsigned32BitPointers, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Floating pointers")]
        [Description("Show float* and double* values as address plus dereferenced value. Applies only when Basic scalar values is enabled.")]
        public bool ShowFloatingPointPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.FloatingPointPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.FloatingPointPointers, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Enum pointers")]
        [Description("Show enum* values as address plus dereferenced enum value. Applies only when Enum values is enabled.")]
        public bool ShowEnumPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.EnumPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.EnumPointers, value);
        }

        [Category("Detailed Type Rules")]
        [DisplayName("Class/struct pointers")]
        [Description("Show class* and struct* values as address plus a short dereferenced summary when possible.")]
        public bool ShowStructPointers
        {
            get => HasTypeRuleKind(InlineValueTypeRuleKinds.StructPointers);
            set => SetTypeRuleKind(InlineValueTypeRuleKinds.StructPointers, value);
        }

        [Category("Rules")]
        [DisplayName("Custom rules")]
        [Description("One rule per line. Syntax: show type:<pattern>, hide type:<pattern>, show name:<pattern>, hide expr:<pattern>. Wildcards * and ? are supported.")]
        [Editor(typeof(MultilineStringEditor), typeof(UITypeEditor))]
        public string CustomRulesText
        {
            get => customRulesText;
            set => customRulesText = value ?? string.Empty;
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
                ruleKinds = settings.RuleKinds;
                typeRuleKinds = settings.TypeRuleKinds;
                customRulesText = settings.CustomRulesText;
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
                settings.SetRuleKinds(ruleKinds);
                settings.SetTypeRuleKinds(typeRuleKinds);
                settings.SetCustomRulesText(customRulesText);
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

        private bool HasRuleKind(InlineValueRuleKinds kind)
        {
            return (ruleKinds & kind) == kind;
        }

        private void SetRuleKind(InlineValueRuleKinds kind, bool enabled)
        {
            if (enabled)
            {
                ruleKinds |= kind;
            }
            else
            {
                ruleKinds &= ~kind;
            }
        }

        private bool HasTypeRuleKind(InlineValueTypeRuleKinds kind)
        {
            return (typeRuleKinds & kind) == kind;
        }

        private void SetTypeRuleKind(InlineValueTypeRuleKinds kind, bool enabled)
        {
            if (enabled)
            {
                typeRuleKinds |= kind;
            }
            else
            {
                typeRuleKinds &= ~kind;
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
