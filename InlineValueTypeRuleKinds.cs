using System;

namespace InlineCppVarDbg
{
    [Flags]
    internal enum InlineValueTypeRuleKinds
    {
        None = 0,
        BooleanScalars = 1 << 0,
        CharacterScalars = 1 << 1,
        IntegralScalars = 1 << 2,
        FloatingPointScalars = 1 << 3,
        EnumValues = 1 << 4,
        BooleanArrays = 1 << 5,
        CharacterArrays = 1 << 6,
        IntegralArrays = 1 << 7,
        FloatingPointArrays = 1 << 8,
        EnumArrays = 1 << 9,
        BooleanPointers = 1 << 10,
        CharacterPointers = 1 << 11,
        IntegralPointers = 1 << 12,
        FloatingPointPointers = 1 << 13,
        EnumPointers = 1 << 14,
        StructPointers = 1 << 15,
        All = BooleanScalars |
              CharacterScalars |
              IntegralScalars |
              FloatingPointScalars |
              EnumValues |
              BooleanArrays |
              CharacterArrays |
              IntegralArrays |
              FloatingPointArrays |
              EnumArrays |
              BooleanPointers |
              CharacterPointers |
              IntegralPointers |
              FloatingPointPointers |
              EnumPointers |
              StructPointers,
    }
}
