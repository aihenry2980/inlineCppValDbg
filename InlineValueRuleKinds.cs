using System;

namespace InlineCppVarDbg
{
    [Flags]
    internal enum InlineValueRuleKinds
    {
        None = 0,
        IntegralPointers = 1 << 0,
        ParameterAnchors = 1 << 1,
        ParseInactivePreprocessor = 1 << 2,
        All = IntegralPointers |
              ParameterAnchors |
              ParseInactivePreprocessor,
    }
}
