using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace BioDraw
{
    internal static class MaterialLibraryService
    {
        private static readonly HashSet<string> MaterialFileExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".svg", ".emf", ".wmf"
        };

        internal static int GetDisplayLength(string text)
        {
            int len = 0;
            foreach (char c in text)
            {
                len += c > 255 ? 2 : 1;
            }
            return len;
        }

        internal static string ToFixedLengthMaterialLabel(string text)
        {
            const int maxVisibleLength = 16;
            const int totalLength = 20;
            const char padChar = ' ';

            var normalized = (text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(normalized))
            {
                return new string(padChar, totalLength);
            }

            string label = "";
            int currentLen = 0;
            bool truncated = false;

            foreach (char c in normalized)
            {
                int charLen = c > 255 ? 2 : 1;
                if (currentLen + charLen > maxVisibleLength)
                {
                    truncated = true;
                    break;
                }
                label += c;
                currentLen += charLen;
            }

            if (truncated)
            {
                label += "…";
                currentLen += 2; // '…' is usually full-width
            }

            int padCount = totalLength - currentLen;
            if (padCount > 0)
            {
                label += new string(padChar, padCount);
            }

            return label;
        }

        internal static string ToEllipsisLabel(string text, int maxWidthPixels, int maxLines)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            if (maxWidthPixels <= 8)
            {
                return "…";
            }

            var sourceText = text.Trim();
            using (var font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point))
            {
                const TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;
                if (MeasureTextWidth(sourceText, font, flags) <= maxWidthPixels)
                {
                    return sourceText;
                }

                var safeLines = Math.Max(1, maxLines);
                if (safeLines >= 2)
                {
                    var firstLineLength = GetLineBreakLength(sourceText, maxWidthPixels, font, flags);
                    if (firstLineLength > 0)
                    {
                        var firstLine = sourceText.Substring(0, firstLineLength).TrimEnd();
                        var remaining = sourceText.Substring(firstLineLength).TrimStart();
                        if (!string.IsNullOrWhiteSpace(remaining))
                        {
                            var secondLine = BuildEllipsisLine(remaining, maxWidthPixels, font, flags);
                            return string.Concat(firstLine, "\n", secondLine);
                        }
                    }
                }

                return BuildEllipsisLine(sourceText, maxWidthPixels, font, flags);
            }
        }

        internal static int MeasureTextWidth(string text, Font font, TextFormatFlags flags)
        {
            return TextRenderer.MeasureText(text, font, new Size(int.MaxValue, int.MaxValue), flags).Width;
        }

        internal static int GetLineBreakLength(string text, int maxWidthPixels, Font font, TextFormatFlags flags)
        {
            var maxLength = GetMaxFittingLength(text, maxWidthPixels, font, flags);
            if (maxLength <= 0)
            {
                return 0;
            }

            var breakLength = maxLength;
            for (int index = maxLength - 1; index >= 1; index--)
            {
                if (char.IsWhiteSpace(text[index]))
                {
                    breakLength = index;
                    break;
                }
            }

            while (breakLength > 0 && char.IsWhiteSpace(text[breakLength - 1]))
            {
                breakLength--;
            }

            return breakLength > 0 ? breakLength : maxLength;
        }

        internal static int GetMaxFittingLength(string text, int maxWidthPixels, Font font, TextFormatFlags flags)
        {
            var low = 1;
            var high = text.Length;
            var best = 0;
            while (low <= high)
            {
                var mid = low + ((high - low) / 2);
                var candidate = text.Substring(0, mid);
                if (MeasureTextWidth(candidate, font, flags) <= maxWidthPixels)
                {
                    best = mid;
                    low = mid + 1;
                }
                else
                {
                    high = mid - 1;
                }
            }

            return best;
        }

        internal static string BuildEllipsisLine(string text, int maxWidthPixels, Font font, TextFormatFlags flags)
        {
            if (MeasureTextWidth(text, font, flags) <= maxWidthPixels)
            {
                return text;
            }

            for (int length = text.Length - 1; length >= 1; length--)
            {
                var candidate = text.Substring(0, length).TrimEnd() + "…";
                if (MeasureTextWidth(candidate, font, flags) <= maxWidthPixels)
                {
                    return candidate;
                }
            }

            return "…";
        }

        internal static int NormalizeIndex(int index, int count)
        {
            if (count <= 0)
            {
                return 0;
            }
            if (index < 0 || index >= count)
            {
                return 0;
            }
            return index;
        }

        internal static bool IsSupportedMaterialFile(string filePath)
        {
            return MaterialFileExtensions.Contains(Path.GetExtension(filePath) ?? string.Empty);
        }

        internal static List<string> GetSubDirectoryNames(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                return new List<string> { "默认" };
            }

            try
            {
                var names = Directory.GetDirectories(folderPath)
                    .Select(Path.GetFileName)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (names.Count > 0)
                {
                    return names;
                }
            }
            catch
            {
            }

            return new List<string> { "默认" };
        }

        internal static List<MaterialEntry> GetMaterialEntriesFromFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                return new List<MaterialEntry> { new MaterialEntry { Name = "默认", FilePath = string.Empty } };
            }

            try
            {
                var entries = Directory.GetFiles(folderPath, "*", SearchOption.TopDirectoryOnly)
                    .Where(IsSupportedMaterialFile)
                    .Select(path => new MaterialEntry
                    {
                        Name = Path.GetFileNameWithoutExtension(path),
                        FilePath = path
                    })
                    .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (entries.Count > 0)
                {
                    return entries;
                }
            }
            catch
            {
            }

            return new List<MaterialEntry> { new MaterialEntry { Name = "默认", FilePath = string.Empty } };
        }

        internal static bool TryBuildMaterialThumbnail(string filePath, string label, int width, int height, out Bitmap bitmap)
        {
            bitmap = null;
            var safeWidth = Math.Max(24, width);
            var safeHeight = Math.Max(24, height);

            try
            {
                using (var image = Image.FromFile(filePath))
                {
                    bitmap = BuildThumbnailBitmap(image, safeWidth, safeHeight, label);
                    return bitmap != null;
                }
            }
            catch
            {
            }

            return TryBuildMaterialThumbnailByPowerPoint(filePath, label, safeWidth, safeHeight, out bitmap);
        }

        internal static Bitmap BuildThumbnailBitmap(Image image, int width, int height, string label)
        {
            var bitmap = new Bitmap(width, height);
            var framePadding = 2f;
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                var frameRect = new RectangleF(0.5f, 0.5f, width - 1f, height - 1f);
                var mediaRect = new RectangleF(
                    framePadding,
                    framePadding,
                    width - (framePadding * 2f),
                    height - (framePadding * 2f));

                using (var pen = new Pen(Color.FromArgb(180, 180, 180), 1f))
                {
                    graphics.FillRectangle(Brushes.White, mediaRect);

                    var scale = Math.Min(mediaRect.Width / Math.Max(1f, image.Width), mediaRect.Height / Math.Max(1f, image.Height));
                    var drawWidth = Math.Max(1f, image.Width * scale);
                    var drawHeight = Math.Max(1f, image.Height * scale);
                    var x = mediaRect.Left + (mediaRect.Width - drawWidth) / 2f;
                    var y = mediaRect.Top + (mediaRect.Height - drawHeight) / 2f;
                    graphics.DrawImage(image, x, y, drawWidth, drawHeight);

                    graphics.DrawRectangle(pen, frameRect.X, frameRect.Y, frameRect.Width, frameRect.Height);
                }
            }

            return bitmap;
        }

        internal static bool TryBuildMaterialThumbnailByPowerPoint(string filePath, string label, int width, int height, out Bitmap bitmap)
        {
            bitmap = null;
            dynamic shape = null;
            string tempPngPath = null;

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    return false;
                }

                dynamic slide = null;
                try
                {
                    slide = app.ActiveWindow?.View?.Slide;
                }
                catch
                {
                }

                if (slide == null)
                {
                    return false;
                }

                shape = slide.Shapes.AddPicture(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    -5000f,
                    -5000f,
                    -1f,
                    -1f);

                tempPngPath = Path.Combine(Path.GetTempPath(), "BioDraw", "material_thumb_" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + ".png");
                Directory.CreateDirectory(Path.GetDirectoryName(tempPngPath));
                shape.Export(tempPngPath, 2);

                using (var image = Image.FromFile(tempPngPath))
                {
                    bitmap = BuildThumbnailBitmap(image, width, height, label);
                }

                return bitmap != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                try
                {
                    if (shape != null)
                    {
                        shape.Delete();
                    }
                }
                catch
                {
                }

                try
                {
                    if (!string.IsNullOrWhiteSpace(tempPngPath) && File.Exists(tempPngPath))
                    {
                        File.Delete(tempPngPath);
                    }
                }
                catch
                {
                }
            }
        }

        internal static bool TryInsertMaterialToCurrentSlide(string filePath, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                errorMessage = "素材文件不存在。";
                return false;
            }

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    errorMessage = "未能获取 PowerPoint 应用实例。";
                    return false;
                }

                dynamic slide = null;
                try
                {
                    slide = app.ActiveWindow?.View?.Slide;
                }
                catch
                {
                }

                if (slide == null)
                {
                    errorMessage = "请先切换到普通编辑视图。";
                    return false;
                }

                dynamic newShape = slide.Shapes.AddPicture(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    0f,
                    0f,
                    -1f,
                    -1f);

                var pageSetup = app.ActivePresentation?.PageSetup;
                if (pageSetup != null)
                {
                    float slideWidth = (float)pageSetup.SlideWidth;
                    float slideHeight = (float)pageSetup.SlideHeight;
                    newShape.Left = (slideWidth - (float)newShape.Width) / 2f;
                    newShape.Top = (slideHeight - (float)newShape.Height) / 2f;
                }

                if (string.Equals(Path.GetExtension(filePath), ".svg", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        newShape.Select();
                        Ribbon1.TryExecuteMso(app, new[] { "SVGEdit" });
                    }
                    catch
                    {
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }
    }
}
