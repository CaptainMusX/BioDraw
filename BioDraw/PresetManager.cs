using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace BioDraw
{
    internal static class PresetManager
    {
        public static ImageReplacePreset CreateDefaultPreset()
        {
            return new ImageReplacePreset
            {
                Name = "默认预设",
                SortOrder = 1,
                FuzzPercent = 5,
                TargetColor = "white",
                Mode = "transparent",
                ReplacementColor = "black"
            };
        }

        public static string GenerateNewPresetName(List<ImageReplacePreset> presets)
        {
            var index = 1;
            while (true)
            {
                var candidate = "新预设" + index.ToString(CultureInfo.InvariantCulture);
                if (presets.All(p => !string.Equals(p.Name, candidate, StringComparison.OrdinalIgnoreCase)))
                {
                    return candidate;
                }
                index++;
            }
        }

        public static int ParseSortOrder(string sortText, int fallbackValue)
        {
            int sortOrder;
            if (int.TryParse(sortText, NumberStyles.Integer, CultureInfo.InvariantCulture, out sortOrder) && sortOrder > 0)
            {
                return sortOrder;
            }
            return Math.Max(1, fallbackValue);
        }

        public static int ParseMaterialPreviewCount(string countText, int maxButtonCount)
        {
            int count;
            if (int.TryParse(countText, NumberStyles.Integer, CultureInfo.InvariantCulture, out count))
            {
                return Math.Max(1, Math.Min(maxButtonCount, count));
            }
            return 5;
        }

        public static double ParseFuzz(string fuzzText)
        {
            double fuzz;
            if (double.TryParse(fuzzText, NumberStyles.Float, CultureInfo.InvariantCulture, out fuzz))
            {
                return NormalizeFuzzPercent(fuzz);
            }
            return 5;
        }

        public static double NormalizeFuzzPercent(double fuzz)
        {
            if (double.IsNaN(fuzz) || double.IsInfinity(fuzz))
            {
                return 5.0;
            }
            if (fuzz < 0)
            {
                fuzz = 0;
            }
            if (fuzz > 100)
            {
                fuzz = 100;
            }
            return Math.Round(fuzz, 1, MidpointRounding.AwayFromZero);
        }

        public static bool ParseBool(string boolText)
        {
            bool value;
            if (bool.TryParse(boolText, out value))
            {
                return value;
            }
            return false;
        }

        public static bool TryParseEditorBounds(XElement root, out System.Drawing.Rectangle bounds)
        {
            bounds = System.Drawing.Rectangle.Empty;
            if (root == null)
            {
                return false;
            }

            int x, y, w, h;
            if (!int.TryParse((string)root.Attribute("EditorX"), NumberStyles.Integer, CultureInfo.InvariantCulture, out x) ||
                !int.TryParse((string)root.Attribute("EditorY"), NumberStyles.Integer, CultureInfo.InvariantCulture, out y) ||
                !int.TryParse((string)root.Attribute("EditorWidth"), NumberStyles.Integer, CultureInfo.InvariantCulture, out w) ||
                !int.TryParse((string)root.Attribute("EditorHeight"), NumberStyles.Integer, CultureInfo.InvariantCulture, out h))
            {
                return false;
            }

            if (w < 620 || h < 360)
            {
                return false;
            }

            bounds = new System.Drawing.Rectangle(x, y, w, h);
            return true;
        }

        public static bool IsLegacyDefaultColorOption(string value)
        {
            return string.Equals(value, "white", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "red", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "blue", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "green", StringComparison.OrdinalIgnoreCase);
        }

        public static void TrimLegacyDefaultColorOptions(List<string> options)
        {
            if (options == null || options.Count == 0)
            {
                return;
            }

            if (!options.All(IsLegacyDefaultColorOption))
            {
                return;
            }

            options.Clear();
            options.Add("white");
        }

        public static string NormalizeColorInputText(string text)
        {
            var normalized = (text ?? string.Empty).Trim();
            if (string.Equals(normalized, "⁠", StringComparison.Ordinal))
            {
                return string.Empty;
            }
            return normalized;
        }

        public static string ToStorageColorInputText(string text)
        {
            var normalized = NormalizeColorInputText(text);
            return string.IsNullOrEmpty(normalized) ? "⁠" : normalized;
        }

        public static bool HasVisibleColorText(string text)
        {
            return !string.IsNullOrWhiteSpace(NormalizeColorInputText(text));
        }

        public static string ToRibbonColorInputText(string text)
        {
            return ToStorageColorInputText(text);
        }

        public static int FindColorOptionIndex(List<string> options, string value)
        {
            if (options == null || options.Count == 0 || string.IsNullOrWhiteSpace(value))
            {
                return -1;
            }

            for (int i = 0; i < options.Count; i++)
            {
                if (string.Equals(options[i], value, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1;
        }

        public static int FindExactMatchIndex(List<string> options, string value)
        {
            return FindColorOptionIndex(options, value);
        }

        public static bool AddColorOption(List<string> options, string value)
        {
            if (options == null)
            {
                return false;
            }

            var normalized = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            if (options.Any(x => string.Equals(x, normalized, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            options.Add(normalized);
            return true;
        }

        public static bool UpsertColorOptionAtPosition(List<string> options, string value, int oneBasedPosition)
        {
            if (options == null)
            {
                return false;
            }

            var normalized = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            var existingIndex = FindColorOptionIndex(options, normalized);
            if (existingIndex >= 0)
            {
                options.RemoveAt(existingIndex);
            }

            var insertIndex = Math.Max(0, Math.Min(options.Count, oneBasedPosition - 1));
            options.Insert(insertIndex, normalized);
            return true;
        }

        public static bool RemoveColorOption(List<string> options, string value)
        {
            if (options == null || options.Count == 0 || string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var index = FindColorOptionIndex(options, value);
            if (index < 0)
            {
                return false;
            }

            options.RemoveAt(index);
            return true;
        }

        public static IEnumerable<ImageReplacePreset> GetPresetsInDisplayOrder(List<ImageReplacePreset> presets)
        {
            return presets.OrderBy(p => p.SortOrder).ThenBy(p => p.Name, StringComparer.OrdinalIgnoreCase);
        }

        public static void NormalizePresetSortOrders(List<ImageReplacePreset> presets)
        {
            if (presets == null || presets.Count == 0)
            {
                return;
            }

            var ordered = GetPresetsInDisplayOrder(presets).ToList();
            for (int i = 0; i < ordered.Count; i++)
            {
                ordered[i].SortOrder = i + 1;
            }
        }
    }
}
