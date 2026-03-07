using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Settings;
using System;

namespace InlineCppVarDbg
{
    internal sealed class InlineValuesSettings
    {
        public const string CollectionPath = "InlineCppVarDbg";
        public const string EnabledPropertyName = "Enabled";
        public const string PreviousLineCountPropertyName = "PreviousLineCount";
        public const string DisplayModePropertyName = "DisplayMode";
        public const string ValueBackgroundColorPropertyName = "ValueBackgroundColor";
        public const string UninitializedValueBackgroundColorPropertyName = "UninitializedValueBackgroundColor";
        public const string ValueChangedAccentColorPropertyName = "ValueChangedAccentColor";
        public const string ValueChipFontSizePropertyName = "ValueChipFontSize";
        public const string NumericDisplayModePropertyName = "NumericDisplayMode";
        private const int DefaultPreviousLineCount = 20;
        private const InlineValueDisplayMode DefaultDisplayMode = InlineValueDisplayMode.Inline;
        private const InlineValueNumericDisplayMode DefaultNumericDisplayMode = InlineValueNumericDisplayMode.Decimal;
        private const int MaxPreviousLineCount = 200;
        private const string DefaultValueBackgroundColor = "#E5E5E5";
        private const string DefaultUninitializedValueBackgroundColor = "#FFF59D";
        private const string DefaultValueChangedAccentColor = "#FF8FB1";
        private const string LegacyDefaultValueChangedAccentColor = "#FF0000";
        private const int DefaultValueChipFontSize = 10;
        private const int MinValueChipFontSize = 8;
        private const int MaxValueChipFontSize = 20;

        private readonly WritableSettingsStore store;
        private readonly object gate = new object();
        private bool enabled;
        private int previousLineCount;
        private InlineValueDisplayMode displayMode;
        private string valueBackgroundColor;
        private string uninitializedValueBackgroundColor;
        private string valueChangedAccentColor;
        private int valueChipFontSize;
        private InlineValueNumericDisplayMode numericDisplayMode;

        public event EventHandler SettingsChanged;

        public InlineValuesSettings(IServiceProvider serviceProvider)
        {
            var manager = new ShellSettingsManager(serviceProvider);
            store = manager.GetWritableSettingsStore(SettingsScope.UserSettings);

            if (!store.CollectionExists(CollectionPath))
            {
                store.CreateCollection(CollectionPath);
            }

            enabled = ReadEnabledOrDefault();
            previousLineCount = ReadPreviousLineCountOrDefault();
            displayMode = ReadDisplayModeOrDefault();
            valueBackgroundColor = ReadValueBackgroundColorOrDefault();
            uninitializedValueBackgroundColor = ReadUninitializedValueBackgroundColorOrDefault();
            valueChangedAccentColor = ReadValueChangedAccentColorOrDefault();
            valueChipFontSize = ReadValueChipFontSizeOrDefault();
            numericDisplayMode = ReadNumericDisplayModeOrDefault();
        }

        public bool IsEnabled
        {
            get
            {
                lock (gate)
                {
                    return enabled;
                }
            }
        }

        public int PreviousLineCount
        {
            get
            {
                lock (gate)
                {
                    return previousLineCount;
                }
            }
        }

        public string ValueBackgroundColor
        {
            get
            {
                lock (gate)
                {
                    return valueBackgroundColor;
                }
            }
        }

        public InlineValueDisplayMode DisplayMode
        {
            get
            {
                lock (gate)
                {
                    return displayMode;
                }
            }
        }

        public string UninitializedValueBackgroundColor
        {
            get
            {
                lock (gate)
                {
                    return uninitializedValueBackgroundColor;
                }
            }
        }

        public string ValueChangedAccentColor
        {
            get
            {
                lock (gate)
                {
                    return valueChangedAccentColor;
                }
            }
        }

        public int ValueChipFontSize
        {
            get
            {
                lock (gate)
                {
                    return valueChipFontSize;
                }
            }
        }

        public InlineValueNumericDisplayMode NumericDisplayMode
        {
            get
            {
                lock (gate)
                {
                    return numericDisplayMode;
                }
            }
        }

        public void SetEnabled(bool newValue)
        {
            bool changed;

            lock (gate)
            {
                if (enabled == newValue)
                {
                    return;
                }

                enabled = newValue;
                PersistEnabled(newValue);
                changed = true;
            }

            if (changed)
            {
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public void SetPreviousLineCount(int value)
        {
            int clamped = ClampPreviousLineCount(value);
            bool changed;

            lock (gate)
            {
                if (previousLineCount == clamped)
                {
                    return;
                }

                previousLineCount = clamped;
                PersistPreviousLineCount(clamped);
                changed = true;
            }

            if (changed)
            {
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public void SetValueBackgroundColor(string color)
        {
            string normalized = NormalizeColorOrDefault(color);
            bool changed;

            lock (gate)
            {
                if (string.Equals(valueBackgroundColor, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                valueBackgroundColor = normalized;
                PersistValueBackgroundColor(normalized);
                changed = true;
            }

            if (changed)
            {
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public void SetDisplayMode(InlineValueDisplayMode mode)
        {
            InlineValueDisplayMode normalized = NormalizeDisplayMode(mode);
            bool changed;

            lock (gate)
            {
                if (displayMode == normalized)
                {
                    return;
                }

                displayMode = normalized;
                PersistDisplayMode(normalized);
                changed = true;
            }

            if (changed)
            {
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public void SetValueChangedAccentColor(string color)
        {
            string normalized = NormalizeColor(color, DefaultValueChangedAccentColor);
            bool changed;

            lock (gate)
            {
                if (string.Equals(valueChangedAccentColor, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                valueChangedAccentColor = normalized;
                PersistValueChangedAccentColor(normalized);
                changed = true;
            }

            if (changed)
            {
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public void SetValueChipFontSize(int size)
        {
            int clamped = ClampValueChipFontSize(size);
            bool changed;

            lock (gate)
            {
                if (valueChipFontSize == clamped)
                {
                    return;
                }

                valueChipFontSize = clamped;
                PersistValueChipFontSize(clamped);
                changed = true;
            }

            if (changed)
            {
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public void SetUninitializedValueBackgroundColor(string color)
        {
            string normalized = NormalizeColor(color, DefaultUninitializedValueBackgroundColor);
            bool changed;

            lock (gate)
            {
                if (string.Equals(uninitializedValueBackgroundColor, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                uninitializedValueBackgroundColor = normalized;
                PersistUninitializedValueBackgroundColor(normalized);
                changed = true;
            }

            if (changed)
            {
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public void SetNumericDisplayMode(InlineValueNumericDisplayMode mode)
        {
            InlineValueNumericDisplayMode normalized = NormalizeNumericDisplayMode(mode);
            bool changed;

            lock (gate)
            {
                if (numericDisplayMode == normalized)
                {
                    return;
                }

                numericDisplayMode = normalized;
                PersistNumericDisplayMode(normalized);
                changed = true;
            }

            if (changed)
            {
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        private bool ReadEnabledOrDefault()
        {
            if (store.PropertyExists(CollectionPath, EnabledPropertyName))
            {
                try
                {
                    return store.GetBoolean(CollectionPath, EnabledPropertyName);
                }
                catch
                {
                    return true;
                }
            }

            PersistEnabled(true);
            return true;
        }

        private int ReadPreviousLineCountOrDefault()
        {
            if (store.PropertyExists(CollectionPath, PreviousLineCountPropertyName))
            {
                try
                {
                    return ClampPreviousLineCount(store.GetInt32(CollectionPath, PreviousLineCountPropertyName));
                }
                catch
                {
                    return DefaultPreviousLineCount;
                }
            }

            PersistPreviousLineCount(DefaultPreviousLineCount);
            return DefaultPreviousLineCount;
        }

        private string ReadValueBackgroundColorOrDefault()
        {
            if (store.PropertyExists(CollectionPath, ValueBackgroundColorPropertyName))
            {
                try
                {
                    return NormalizeColorOrDefault(store.GetString(CollectionPath, ValueBackgroundColorPropertyName));
                }
                catch
                {
                    return DefaultValueBackgroundColor;
                }
            }

            PersistValueBackgroundColor(DefaultValueBackgroundColor);
            return DefaultValueBackgroundColor;
        }

        private InlineValueDisplayMode ReadDisplayModeOrDefault()
        {
            if (store.PropertyExists(CollectionPath, DisplayModePropertyName))
            {
                try
                {
                    return NormalizeDisplayMode((InlineValueDisplayMode)store.GetInt32(CollectionPath, DisplayModePropertyName));
                }
                catch
                {
                    return DefaultDisplayMode;
                }
            }

            PersistDisplayMode(DefaultDisplayMode);
            return DefaultDisplayMode;
        }

        private string ReadUninitializedValueBackgroundColorOrDefault()
        {
            if (store.PropertyExists(CollectionPath, UninitializedValueBackgroundColorPropertyName))
            {
                try
                {
                    return NormalizeColor(store.GetString(CollectionPath, UninitializedValueBackgroundColorPropertyName), DefaultUninitializedValueBackgroundColor);
                }
                catch
                {
                    return DefaultUninitializedValueBackgroundColor;
                }
            }

            PersistUninitializedValueBackgroundColor(DefaultUninitializedValueBackgroundColor);
            return DefaultUninitializedValueBackgroundColor;
        }

        private string ReadValueChangedAccentColorOrDefault()
        {
            if (store.PropertyExists(CollectionPath, ValueChangedAccentColorPropertyName))
            {
                try
                {
                    string normalized = NormalizeColor(store.GetString(CollectionPath, ValueChangedAccentColorPropertyName), DefaultValueChangedAccentColor);
                    if (string.Equals(normalized, LegacyDefaultValueChangedAccentColor, StringComparison.OrdinalIgnoreCase))
                    {
                        PersistValueChangedAccentColor(DefaultValueChangedAccentColor);
                        return DefaultValueChangedAccentColor;
                    }

                    return normalized;
                }
                catch
                {
                    return DefaultValueChangedAccentColor;
                }
            }

            PersistValueChangedAccentColor(DefaultValueChangedAccentColor);
            return DefaultValueChangedAccentColor;
        }

        private int ReadValueChipFontSizeOrDefault()
        {
            if (store.PropertyExists(CollectionPath, ValueChipFontSizePropertyName))
            {
                try
                {
                    return ClampValueChipFontSize(store.GetInt32(CollectionPath, ValueChipFontSizePropertyName));
                }
                catch
                {
                    return DefaultValueChipFontSize;
                }
            }

            PersistValueChipFontSize(DefaultValueChipFontSize);
            return DefaultValueChipFontSize;
        }

        private InlineValueNumericDisplayMode ReadNumericDisplayModeOrDefault()
        {
            if (store.PropertyExists(CollectionPath, NumericDisplayModePropertyName))
            {
                try
                {
                    return NormalizeNumericDisplayMode((InlineValueNumericDisplayMode)store.GetInt32(CollectionPath, NumericDisplayModePropertyName));
                }
                catch
                {
                    return DefaultNumericDisplayMode;
                }
            }

            PersistNumericDisplayMode(DefaultNumericDisplayMode);
            return DefaultNumericDisplayMode;
        }

        private void PersistEnabled(bool value)
        {
            store.SetBoolean(CollectionPath, EnabledPropertyName, value);
        }

        private void PersistPreviousLineCount(int value)
        {
            store.SetInt32(CollectionPath, PreviousLineCountPropertyName, ClampPreviousLineCount(value));
        }

        private void PersistValueBackgroundColor(string value)
        {
            store.SetString(CollectionPath, ValueBackgroundColorPropertyName, NormalizeColorOrDefault(value));
        }

        private void PersistDisplayMode(InlineValueDisplayMode mode)
        {
            store.SetInt32(CollectionPath, DisplayModePropertyName, (int)NormalizeDisplayMode(mode));
        }

        private void PersistUninitializedValueBackgroundColor(string value)
        {
            store.SetString(CollectionPath, UninitializedValueBackgroundColorPropertyName, NormalizeColor(value, DefaultUninitializedValueBackgroundColor));
        }

        private void PersistValueChangedAccentColor(string value)
        {
            store.SetString(CollectionPath, ValueChangedAccentColorPropertyName, NormalizeColor(value, DefaultValueChangedAccentColor));
        }

        private void PersistValueChipFontSize(int size)
        {
            store.SetInt32(CollectionPath, ValueChipFontSizePropertyName, ClampValueChipFontSize(size));
        }

        private void PersistNumericDisplayMode(InlineValueNumericDisplayMode mode)
        {
            store.SetInt32(CollectionPath, NumericDisplayModePropertyName, (int)NormalizeNumericDisplayMode(mode));
        }

        private static int ClampPreviousLineCount(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            if (value > MaxPreviousLineCount)
            {
                return MaxPreviousLineCount;
            }

            return value;
        }

        private static int ClampValueChipFontSize(int value)
        {
            if (value < MinValueChipFontSize)
            {
                return MinValueChipFontSize;
            }

            if (value > MaxValueChipFontSize)
            {
                return MaxValueChipFontSize;
            }

            return value;
        }

        private static InlineValueDisplayMode NormalizeDisplayMode(InlineValueDisplayMode mode)
        {
            if (mode != InlineValueDisplayMode.Inline && mode != InlineValueDisplayMode.EndOfLine)
            {
                return DefaultDisplayMode;
            }

            return mode;
        }

        private static InlineValueNumericDisplayMode NormalizeNumericDisplayMode(InlineValueNumericDisplayMode mode)
        {
            if (mode != InlineValueNumericDisplayMode.Decimal &&
                mode != InlineValueNumericDisplayMode.Hexadecimal &&
                mode != InlineValueNumericDisplayMode.Binary)
            {
                return DefaultNumericDisplayMode;
            }

            return mode;
        }

        private static string NormalizeColorOrDefault(string value)
        {
            return NormalizeColor(value, DefaultValueBackgroundColor);
        }

        private static string NormalizeColor(string value, string fallback)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            string trimmed = value.Trim();
            if (trimmed.Length == 7 && trimmed[0] == '#')
            {
                if (IsHex(trimmed, 1, 6))
                {
                    return trimmed.ToUpperInvariant();
                }
            }

            if (trimmed.Length == 9 && trimmed[0] == '#')
            {
                if (IsHex(trimmed, 1, 8))
                {
                    return trimmed.ToUpperInvariant();
                }
            }

            return fallback;
        }

        private static bool IsHex(string value, int start, int count)
        {
            int end = start + count;
            for (int i = start; i < end; i++)
            {
                char ch = value[i];
                bool isHex =
                    (ch >= '0' && ch <= '9') ||
                    (ch >= 'a' && ch <= 'f') ||
                    (ch >= 'A' && ch <= 'F');
                if (!isHex)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
