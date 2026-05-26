using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace BioDraw
{
    internal static class ImageReplacePipeline
    {
        internal static bool RunImageMagickReplace(string imageMagickPath, string sourcePath, string outputPath, ImageReplacePreset preset, out string errorMessage)
        {
            errorMessage = string.Empty;

            var arguments = new StringBuilder();
            arguments.Append(QuoteArg(sourcePath));
            arguments.Append(" -fuzz ");
            arguments.Append(preset.FuzzPercent.ToString("0.##", CultureInfo.InvariantCulture));
            arguments.Append("% ");

            if (string.Equals(preset.Mode, "fill", StringComparison.OrdinalIgnoreCase))
            {
                arguments.Append("-fill ");
                arguments.Append(QuoteArg(preset.ReplacementColor));
                arguments.Append(" -opaque ");
                arguments.Append(QuoteArg(preset.TargetColor));
                arguments.Append(" ");
            }
            else
            {
                arguments.Append("-transparent ");
                arguments.Append(QuoteArg(preset.TargetColor));
                arguments.Append(" ");
            }

            var outputExtension = Path.GetExtension(outputPath);
            if (IsJpegFormat(outputExtension))
            {
                arguments.Append("-quality 100 -sampling-factor 4:4:4 -interlace none ");
            }

            arguments.Append(QuoteArg(outputPath));

            try
            {
                var magickFileName = string.IsNullOrWhiteSpace(imageMagickPath) ? "magick" : imageMagickPath;
                var processStartInfo = new ProcessStartInfo
                {
                    FileName = magickFileName,
                    Arguments = arguments.ToString(),
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(processStartInfo))
                {
                    if (process == null)
                    {
                        errorMessage = "无法启动 ImageMagick。";
                        return false;
                    }

                    var stdOut = process.StandardOutput.ReadToEnd();
                    var stdErr = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        errorMessage = string.IsNullOrWhiteSpace(stdErr) ? stdOut : stdErr;
                        if (string.IsNullOrWhiteSpace(errorMessage))
                        {
                            errorMessage = "ImageMagick 执行失败。";
                        }
                        return false;
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

        internal static string QuoteArg(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\"", "\\\"") + "\"";
        }

        internal static bool TryGetSelectedShapes(dynamic selection, out List<dynamic> shapes)
        {
            shapes = new List<dynamic>();
            try
            {
                var shapeRange = selection.ShapeRange;
                if (shapeRange == null)
                {
                    return false;
                }

                var count = 0;
                try
                {
                    count = (int)shapeRange.Count;
                }
                catch
                {
                }

                if (count <= 0)
                {
                    return false;
                }

                for (int i = 1; i <= count; i++)
                {
                    shapes.Add(shapeRange[i]);
                }

                return shapes.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        internal static List<dynamic> FlattenToPictures(List<dynamic> shapes)
        {
            var result = new List<dynamic>();
            foreach (var shape in shapes)
            {
                FlattenToPicturesRecursive(shape, result);
            }
            return result;
        }

        internal static void FlattenToPicturesRecursive(dynamic shape, List<dynamic> result)
        {
            try
            {
                int type = (int)shape.Type;
                if (type == 13)
                {
                    result.Add(shape);
                }
                else if (type == 6)
                {
                    dynamic groupItems = shape.GroupItems;
                    int count = (int)groupItems.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        FlattenToPicturesRecursive(groupItems[i], result);
                    }
                }
            }
            catch
            {
            }
        }

        internal static bool IsShapeInGroup(dynamic shape)
        {
            try
            {
                dynamic parent = shape.Parent;
                return parent != null && (int)parent.Type == 6;
            }
            catch
            {
                return false;
            }
        }

        internal static bool TryReplaceShapePictureWithMagick(dynamic shape, string imageMagickPath, ImageReplacePreset preset, string snapshotPath, out dynamic replacedShape, out string errorMessage)
        {
            replacedShape = null;
            errorMessage = string.Empty;
            if (shape == null)
            {
                errorMessage = "未找到可处理的图片。";
                return false;
            }

            try
            {
                var type = 0;
                try
                {
                    type = (int)shape.Type;
                }
                catch
                {
                }

                if (type != 13)
                {
                    errorMessage = "选中对象不是图片。";
                    return false;
                }

                if (IsShapeInGroup(shape))
                {
                    errorMessage = "组内图片暂不支持图色替换，请先取消组合。";
                    return false;
                }

                string sourcePath;
                if (!TryExtractOriginalImageFromPptx(shape, snapshotPath, out sourcePath, out errorMessage))
                {
                    return false;
                }

                var outputPath = BuildMagickOutputPath(sourcePath, preset);

                if (!RunImageMagickReplace(imageMagickPath, sourcePath, outputPath, preset, out errorMessage))
                {
                    TryDeleteFile(sourcePath);
                    return false;
                }

                if (!TryReplaceShapeImageInPlace(shape, outputPath, out replacedShape, out errorMessage))
                {
                    TryDeleteFile(sourcePath);
                    TryDeleteFile(outputPath);
                    return false;
                }

                TryDeleteFile(sourcePath);
                TryDeleteFile(outputPath);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        internal static string BuildMagickOutputPath(string sourcePath, ImageReplacePreset preset)
        {
            var outputExtension = Path.GetExtension(sourcePath);
            if (string.IsNullOrWhiteSpace(outputExtension))
            {
                outputExtension = ".png";
            }

            if (string.Equals(preset.Mode, "transparent", StringComparison.OrdinalIgnoreCase) &&
                IsTransparencyUnsupportedFormat(outputExtension))
            {
                outputExtension = ".png";
            }

            return Path.Combine(
                Path.GetDirectoryName(sourcePath) ?? Path.GetTempPath(),
                Path.GetFileNameWithoutExtension(sourcePath) + "_magick" + outputExtension);
        }

        internal static bool IsTransparencyUnsupportedFormat(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
            {
                return true;
            }

            switch (extension.Trim().ToLowerInvariant())
            {
                case ".jpg":
                case ".jpeg":
                case ".jpe":
                case ".jfif":
                case ".bmp":
                case ".dib":
                    return true;
                default:
                    return false;
            }
        }

        internal static bool IsJpegFormat(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
            {
                return false;
            }

            switch (extension.Trim().ToLowerInvariant())
            {
                case ".jpg":
                case ".jpeg":
                case ".jpe":
                case ".jfif":
                    return true;
                default:
                    return false;
            }
        }

        internal static bool TryGetPictureCropValues(dynamic shape, out float cropLeft, out float cropTop, out float cropRight, out float cropBottom)
        {
            cropLeft = 0f;
            cropTop = 0f;
            cropRight = 0f;
            cropBottom = 0f;
            try
            {
                cropLeft = (float)shape.PictureFormat.CropLeft;
                cropTop = (float)shape.PictureFormat.CropTop;
                cropRight = (float)shape.PictureFormat.CropRight;
                cropBottom = (float)shape.PictureFormat.CropBottom;
                return true;
            }
            catch
            {
                return false;
            }
        }

        internal static bool TryExtractOriginalImageFromPptx(dynamic shape, string snapshotPath, out string filePath, out string errorMessage)
        {
            filePath = null;
            errorMessage = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
                {
                    errorMessage = "无法读取当前演示文稿快照。";
                    return false;
                }

                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw");
                Directory.CreateDirectory(tempDir);

                var slideIndex = (int)shape.Parent.SlideIndex;
                var shapeId = (int)shape.Id;

                using (var stream = new FileStream(snapshotPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, false))
                {
                    string slidePartPath;
                    if (!TryResolveSlidePartPath(archive, slideIndex, out slidePartPath, out errorMessage))
                    {
                        return false;
                    }

                    string mediaPartPath;
                    if (!TryResolveMediaPathForShape(archive, slidePartPath, shapeId, out mediaPartPath, out errorMessage))
                    {
                        return false;
                    }

                    var mediaEntry = archive.GetEntry(mediaPartPath);
                    if (mediaEntry == null)
                    {
                        errorMessage = "未找到选中图片对应的媒体文件。";
                        return false;
                    }

                    var ext = Path.GetExtension(mediaPartPath);
                    if (string.IsNullOrWhiteSpace(ext))
                    {
                        ext = ".png";
                    }

                    var targetFilePath = Path.Combine(tempDir, "ppt_media_" + Guid.NewGuid().ToString("N") + ext);
                    using (var entryStream = mediaEntry.Open())
                    using (var output = new FileStream(targetFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        entryStream.CopyTo(output);
                    }

                    if (!File.Exists(targetFilePath) || new FileInfo(targetFilePath).Length <= 0)
                    {
                        errorMessage = "读取原始图片失败。";
                        return false;
                    }

                    filePath = targetFilePath;
                    return true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        internal static bool TryCreatePresentationSnapshot(dynamic application, out string snapshotPath, out string errorMessage)
        {
            snapshotPath = null;
            errorMessage = string.Empty;
            try
            {
                var presentation = application?.ActivePresentation;
                if (presentation == null)
                {
                    errorMessage = "未找到当前演示文稿。";
                    return false;
                }

                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw");
                Directory.CreateDirectory(tempDir);
                snapshotPath = Path.Combine(tempDir, "ppt_snapshot_" + Guid.NewGuid().ToString("N") + ".pptx");

                presentation.SaveCopyAs(snapshotPath);
                if (!File.Exists(snapshotPath) || new FileInfo(snapshotPath).Length <= 0)
                {
                    errorMessage = "无法创建演示文稿快照。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "无法读取当前演示文稿，请确认文档已正常打开。 " + ex.Message;
                snapshotPath = null;
                return false;
            }
        }

        internal static void SetStatusText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }
            try
            {
                System.Diagnostics.Debug.WriteLine("[BioDraw] " + text);
            }
            catch
            {
            }
        }

        internal static bool TryResolveSlidePartPath(ZipArchive archive, int slideIndex, out string slidePartPath, out string errorMessage)
        {
            slidePartPath = null;
            errorMessage = string.Empty;
            try
            {
                var presentationEntry = archive.GetEntry("ppt/presentation.xml");
                var presentationRelsEntry = archive.GetEntry("ppt/_rels/presentation.xml.rels");
                if (presentationEntry == null || presentationRelsEntry == null)
                {
                    errorMessage = "PPTX 结构异常，找不到演示文稿索引。";
                    return false;
                }

                XDocument presentationDoc;
                XDocument relsDoc;
                using (var stream = presentationEntry.Open())
                {
                    presentationDoc = XDocument.Load(stream);
                }
                using (var stream = presentationRelsEntry.Open())
                {
                    relsDoc = XDocument.Load(stream);
                }

                var p = (XNamespace)"http://schemas.openxmlformats.org/presentationml/2006/main";
                var r = (XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                var slideIdNodes = presentationDoc.Descendants(p + "sldId").ToList();
                if (slideIndex <= 0 || slideIndex > slideIdNodes.Count)
                {
                    errorMessage = "无法定位选中图片所在幻灯片。";
                    return false;
                }

                var slideRid = (string)slideIdNodes[slideIndex - 1].Attribute(r + "id");
                if (string.IsNullOrWhiteSpace(slideRid))
                {
                    errorMessage = "幻灯片关系索引缺失。";
                    return false;
                }

                var rel = relsDoc.Root?
                    .Elements()
                    .FirstOrDefault(x => string.Equals((string)x.Attribute("Id"), slideRid, StringComparison.Ordinal));
                var target = (string)rel?.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target))
                {
                    errorMessage = "幻灯片关系目标缺失。";
                    return false;
                }

                slidePartPath = ResolveZipPartPath("ppt/presentation.xml", target);
                if (archive.GetEntry(slidePartPath) == null)
                {
                    errorMessage = "未找到幻灯片数据。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        internal static bool TryResolveMediaPathForShape(ZipArchive archive, string slidePartPath, int shapeId, out string mediaPartPath, out string errorMessage)
        {
            mediaPartPath = null;
            errorMessage = string.Empty;
            try
            {
                var slideEntry = archive.GetEntry(slidePartPath);
                var slideRelsPath = GetRelationshipPartPath(slidePartPath);
                var slideRelsEntry = archive.GetEntry(slideRelsPath);
                if (slideEntry == null || slideRelsEntry == null)
                {
                    errorMessage = "找不到幻灯片图片关系文件。";
                    return false;
                }

                XDocument slideDoc;
                XDocument relsDoc;
                using (var stream = slideEntry.Open())
                {
                    slideDoc = XDocument.Load(stream);
                }
                using (var stream = slideRelsEntry.Open())
                {
                    relsDoc = XDocument.Load(stream);
                }

                var p = (XNamespace)"http://schemas.openxmlformats.org/presentationml/2006/main";
                var a = (XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
                var r = (XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                var targetPic = slideDoc.Descendants(p + "pic")
                    .FirstOrDefault(pic =>
                    {
                        var idAttr = (string)pic
                            .Element(p + "nvPicPr")?
                            .Element(p + "cNvPr")?
                            .Attribute("id");
                        int idValue;
                        return int.TryParse(idAttr, out idValue) && idValue == shapeId;
                    });

                if (targetPic == null)
                {
                    errorMessage = "无法定位选中图片对应的原始资源。";
                    return false;
                }

                var embedRid = (string)targetPic
                    .Element(p + "blipFill")?
                    .Element(a + "blip")?
                    .Attribute(r + "embed");
                if (string.IsNullOrWhiteSpace(embedRid))
                {
                    errorMessage = "该图片不包含可提取的嵌入资源。";
                    return false;
                }

                var rel = relsDoc.Root?
                    .Elements()
                    .FirstOrDefault(x => string.Equals((string)x.Attribute("Id"), embedRid, StringComparison.Ordinal));
                var target = (string)rel?.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target))
                {
                    errorMessage = "未找到图片关系映射。";
                    return false;
                }

                mediaPartPath = ResolveZipPartPath(slidePartPath, target);
                if (archive.GetEntry(mediaPartPath) == null)
                {
                    errorMessage = "媒体文件不存在。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        internal static string ResolveZipPartPath(string basePartPath, string relativeTarget)
        {
            var normalizedBase = basePartPath.Replace("\\", "/");
            var normalizedTarget = relativeTarget.Replace("\\", "/");

            if (normalizedTarget.StartsWith("/", StringComparison.Ordinal))
            {
                return normalizedTarget.TrimStart('/');
            }

            var baseUri = new Uri("http://local/" + normalizedBase, UriKind.Absolute);
            var resolvedUri = new Uri(baseUri, normalizedTarget);
            return resolvedUri.AbsolutePath.TrimStart('/');
        }

        internal static string GetRelationshipPartPath(string partPath)
        {
            var normalized = partPath.Replace("\\", "/");
            var lastSlash = normalized.LastIndexOf('/');
            if (lastSlash < 0)
            {
                return "_rels/" + normalized + ".rels";
            }

            var dir = normalized.Substring(0, lastSlash);
            var file = normalized.Substring(lastSlash + 1);
            return dir + "/_rels/" + file + ".rels";
        }

        internal static List<CapturedEffect> CaptureShapeEffects(dynamic shape)
        {
            var captured = new List<CapturedEffect>();
            try
            {
                int shapeId = (int)shape.Id;
                dynamic slide = shape.Parent;
                dynamic timeline = slide.TimeLine;

                CaptureEffectsFromSequence(timeline.MainSequence, shapeId, captured);

                try
                {
                    dynamic interactiveSeqs = timeline.InteractiveSequences;
                    int seqCount = (int)interactiveSeqs.Count;
                    for (int j = 1; j <= seqCount; j++)
                    {
                        CaptureEffectsFromSequence(interactiveSeqs[j], shapeId, captured);
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }
            return captured;
        }

        internal static void CaptureEffectsFromSequence(dynamic sequence, int shapeId, List<CapturedEffect> captured)
        {
            try
            {
                int count = (int)sequence.Count;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic effect = sequence[i];
                        dynamic effectShape = effect.Shape;
                        if (effectShape != null && (int)effectShape.Id == shapeId)
                        {
                            captured.Add(new CapturedEffect
                            {
                                EffectType = (int)effect.EffectType,
                                Exit = (bool)effect.Exit,
                                Duration = (float)effect.Timing.Duration,
                                TriggerType = (int)effect.Timing.TriggerType,
                                TriggerDelayTime = (float)effect.Timing.TriggerDelayTime,
                            });
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        internal static void RestoreShapeEffects(dynamic shape, List<CapturedEffect> effects)
        {
            if (effects == null || effects.Count == 0)
            {
                return;
            }
            try
            {
                dynamic slide = shape.Parent;
                dynamic timeline = slide.TimeLine;
                dynamic mainSeq = timeline.MainSequence;

                foreach (var ce in effects)
                {
                    try
                    {
                        dynamic newEffect = mainSeq.AddEffect(
                            shape,
                            ce.EffectType,
                            0,
                            ce.TriggerType,
                            -1);

                        newEffect.Exit = ce.Exit;
                        newEffect.Timing.Duration = ce.Duration;
                        newEffect.Timing.TriggerDelayTime = ce.TriggerDelayTime;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        internal static bool TryReplaceShapeImageInPlace(dynamic shape, string outputPath, out dynamic newShape, out string errorMessage)
        {
            newShape = null;
            errorMessage = string.Empty;
            try
            {
                if (!File.Exists(outputPath))
                {
                    errorMessage = "输出文件不存在。";
                    return false;
                }

                var left = (float)shape.Left;
                var top = (float)shape.Top;
                var width = (float)shape.Width;
                var height = (float)shape.Height;
                var rotation = (float)shape.Rotation;
                var zOrderPosition = (int)shape.ZOrderPosition;
                var shapeName = string.Empty;
                try
                {
                    shapeName = (string)shape.Name;
                }
                catch
                {
                }
                var lockAspectRatio = 0;
                try
                {
                    lockAspectRatio = (int)shape.LockAspectRatio;
                }
                catch
                {
                }

                var altText = string.Empty;
                try
                {
                    altText = (string)shape.AlternativeText;
                }
                catch
                {
                }
                if (string.IsNullOrWhiteSpace(altText))
                {
                    try
                    {
                        altText = (string)shape.Title;
                    }
                    catch
                    {
                    }
                }

                string hyperlinkAddress = null;
                string hyperlinkSubAddress = null;
                try
                {
                    dynamic clickAction = shape.ActionSettings[1];
                    if (clickAction != null)
                    {
                        int action = (int)clickAction.Action;
                        if (action == 7)
                        {
                            dynamic hl = clickAction.Hyperlink;
                            hyperlinkAddress = (string)hl.Address;
                            hyperlinkSubAddress = (string)hl.SubAddress;
                        }
                    }
                }
                catch
                {
                }

                var cropLeft = 0f;
                var cropTop = 0f;
                var cropRight = 0f;
                var cropBottom = 0f;
                var hasCrop = TryGetPictureCropValues(shape, out cropLeft, out cropTop, out cropRight, out cropBottom) &&
                    (Math.Abs(cropLeft) > 0.01f || Math.Abs(cropTop) > 0.01f || Math.Abs(cropRight) > 0.01f || Math.Abs(cropBottom) > 0.01f);

                var insertionLeft = left;
                var insertionTop = top;
                var insertionWidth = width;
                var insertionHeight = height;
                if (hasCrop)
                {
                    insertionLeft = left - cropLeft;
                    insertionTop = top - cropTop;
                    insertionWidth = Math.Max(1f, width + cropLeft + cropRight);
                    insertionHeight = Math.Max(1f, height + cropTop + cropBottom);
                }

                var capturedEffects = CaptureShapeEffects(shape);

                try
                {
                    shape.PickUp();
                }
                catch
                {
                }

                try
                {
                    var shapes = shape.Parent.Shapes;
                    shape.Delete();

                    newShape = shapes.AddPicture(
                        outputPath,
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue,
                        insertionLeft,
                        insertionTop,
                        insertionWidth,
                        insertionHeight);

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(shapeName))
                        {
                            newShape.Name = shapeName;
                        }
                    }
                    catch
                    {
                    }

                    try
                    {
                        newShape.Rotation = rotation;
                    }
                    catch
                    {
                    }
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(altText))
                        {
                            newShape.AlternativeText = altText;
                        }
                    }
                    catch
                    {
                    }

                    if (hyperlinkAddress != null)
                    {
                        try
                        {
                            dynamic clickAction = newShape.ActionSettings[1];
                            clickAction.Action = 7;
                            dynamic hl = clickAction.Hyperlink;
                            hl.Address = hyperlinkAddress;
                            if (hyperlinkSubAddress != null)
                            {
                                hl.SubAddress = hyperlinkSubAddress;
                            }
                        }
                        catch
                        {
                        }
                    }

                    try
                    {
                        newShape.LockAspectRatio = 0;
                    }
                    catch
                    {
                    }

                    if (hasCrop)
                    {
                        try
                        {
                            newShape.PictureFormat.CropLeft = cropLeft;
                            newShape.PictureFormat.CropTop = cropTop;
                            newShape.PictureFormat.CropRight = cropRight;
                            newShape.PictureFormat.CropBottom = cropBottom;
                        }
                        catch
                        {
                        }
                    }

                    try
                    {
                        newShape.Left = left;
                        newShape.Top = top;
                        newShape.Width = width;
                        newShape.Height = height;
                    }
                    catch
                    {
                    }

                    try
                    {
                        newShape.LockAspectRatio = lockAspectRatio;
                    }
                    catch
                    {
                    }

                    try
                    {
                        newShape.Apply();
                    }
                    catch
                    {
                    }

                    try
                    {
                        RestoreShapeEffects(newShape, capturedEffects);
                    }
                    catch
                    {
                    }

                    try
                    {
                        var guard = 0;
                        var currentZOrder = (int)newShape.ZOrderPosition;
                        while (currentZOrder != zOrderPosition && guard < 2048)
                        {
                            if (currentZOrder > zOrderPosition)
                            {
                                newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward);
                            }
                            else
                            {
                                newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringForward);
                            }
                            currentZOrder = (int)newShape.ZOrderPosition;
                            guard++;
                        }
                    }
                    catch
                    {
                    }
                }
                catch
                {
                    errorMessage = "替换图片失败。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                newShape = null;
                return false;
            }
        }

        internal static void TryReselectShapes(List<dynamic> shapes)
        {
            if (shapes == null || shapes.Count == 0)
            {
                return;
            }

            try
            {
                if (shapes.Count == 1)
                {
                    shapes[0].Select();
                    return;
                }

                var selectedCount = 0;
                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    if (shape == null)
                    {
                        continue;
                    }

                    if (selectedCount == 0)
                    {
                        shape.Select(Microsoft.Office.Core.MsoTriState.msoTrue);
                    }
                    else
                    {
                        shape.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
                    }
                    selectedCount++;
                }

                if (selectedCount > 1)
                {
                    return;
                }

                if (selectedCount == 1)
                {
                    return;
                }

                var parentShapes = shapes[0].Parent.Shapes;
                var ids = new int[shapes.Count];
                for (int i = 0; i < shapes.Count; i++)
                {
                    ids[i] = (int)shapes[i].Id;
                }
                var range = parentShapes.Range(ids);
                range.Select();
            }
            catch
            {
                try
                {
                    shapes[0].Select();
                }
                catch
                {
                }
            }
        }

        internal static void TryDeleteFile(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }
    }

    internal sealed class CapturedEffect
    {
        public int EffectType { get; set; }
        public bool Exit { get; set; }
        public float Duration { get; set; }
        public int TriggerType { get; set; }
        public float TriggerDelayTime { get; set; }
    }
}
