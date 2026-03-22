using System;

namespace InlineCppVarDbg
{
    [Flags]
    internal enum InlineValueEvaluationKinds
    {
        None = 0,
        BasicScalars = 1 << 0,
        Enums = 1 << 1,
        Arrays = 1 << 2,
        GetterCalls = 1 << 3,
        IndexedExpressions = 1 << 4,
        FallbackExpressions = 1 << 5,
        NullPointers = 1 << 6,
        UninitializedValues = 1 << 7,
        All = BasicScalars |
              Enums |
              Arrays |
              GetterCalls |
              IndexedExpressions |
              FallbackExpressions |
              NullPointers |
              UninitializedValues,
    }
}
