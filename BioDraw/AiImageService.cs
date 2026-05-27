using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml.Linq;

namespace BioDraw
{
    internal static class AiImageService
    {
        private const string SettingsRootElement = "BioDrawAiImageSettings";
        public const int AiModelButtonCount = 12;
        public const int DefaultModelPreviewCount = 5;

        public static AiImageApiSettings CreateDefaultSettings()
        {
            return new AiImageApiSettings
            {
                DisplayName = "GPT-Image-2",
                EndpointUrl = "https://www.packyapi.com/v1/images/generations",
                ApiToken = string.Empty,
                Model = "gpt-image-2",
                DefaultWidth = 1024,
                DefaultHeight = 1024,
                DefaultQuality = "auto",
                DefaultFormat = "png",
                IconPath = string.Empty,
                LockAspectRatio = false
            };
        }

        public static AiImageGlobalSettings CreateDefaultGlobalSettings()
        {
            return new AiImageGlobalSettings
            {
                OverridePerModel = false,
                ModelPreviewCount = DefaultModelPreviewCount,
                DefaultWidth = 1024,
                DefaultHeight = 1024,
                DefaultQuality = "auto",
                DefaultFormat = "png"
            };
        }

        public static List<AiImageApiSettings> LoadModelSettings()
        {
            var results = new List<AiImageApiSettings>();
            try
            {
                var path = GetSettingsFilePath();
                if (!File.Exists(path))
                {
                    var defaults = CreateDefaultSettings();
                    results.Add(defaults);
                    return results;
                }

                var doc = XDocument.Load(path);
                var root = doc.Root;
                if (root == null || root.Name.LocalName != SettingsRootElement)
                {
                    var defaults = CreateDefaultSettings();
                    results.Add(defaults);
                    return results;
                }

                foreach (var entry in root.Elements("AiImageApi"))
                {
                    var settings = new AiImageApiSettings
                    {
                        DisplayName = (string)entry.Element("DisplayName") ?? "GPT-Image-2",
                        EndpointUrl = (string)entry.Element("EndpointUrl") ?? "https://www.packyapi.com/v1/images/generations",
                        ApiToken = (string)entry.Element("ApiToken") ?? string.Empty,
                        Model = (string)entry.Element("Model") ?? "gpt-image-2",
                        DefaultQuality = (string)entry.Element("DefaultQuality") ?? "auto",
                        DefaultFormat = (string)entry.Element("DefaultFormat") ?? "png",
                        IconPath = (string)entry.Element("IconPath") ?? string.Empty,
                        LockAspectRatio = ParseBool((string)entry.Element("LockAspectRatio"))
                    };

                    // Width/Height: try new fields first, fall back to legacy DefaultSize
                    var wEl = entry.Element("DefaultWidth");
                    var hEl = entry.Element("DefaultHeight");
                    if (wEl != null && hEl != null)
                    {
                        settings.DefaultWidth = ParseInt((string)wEl, 1024);
                        settings.DefaultHeight = ParseInt((string)hEl, 1024);
                    }
                    else
                    {
                        var sizeStr = (string)entry.Element("DefaultSize");
                        ParseLegacySize(sizeStr, out int w, out int h);
                        settings.DefaultWidth = w;
                        settings.DefaultHeight = h;
                    }

                    results.Add(settings);
                }

                if (results.Count == 0)
                {
                    var defaults = CreateDefaultSettings();
                    results.Add(defaults);
                }

                return results;
            }
            catch
            {
                return new List<AiImageApiSettings> { CreateDefaultSettings() };
            }
        }

        public static AiImageGlobalSettings LoadGlobalSettings()
        {
            var defaults = CreateDefaultGlobalSettings();
            try
            {
                var path = GetSettingsFilePath();
                if (!File.Exists(path))
                    return defaults;

                var doc = XDocument.Load(path);
                var root = doc.Root;
                if (root == null || root.Name.LocalName != SettingsRootElement)
                    return defaults;

                return new AiImageGlobalSettings
                {
                    OverridePerModel = ParseBool((string)root.Attribute("OverridePerModel")),
                    ModelPreviewCount = ClampModelPreviewCount(ParseInt((string)root.Attribute("ModelPreviewCount"), DefaultModelPreviewCount)),
                    DefaultWidth = ParseInt((string)root.Attribute("GlobalWidth"), 1024),
                    DefaultHeight = ParseInt((string)root.Attribute("GlobalHeight"), 1024),
                    DefaultQuality = (string)root.Attribute("GlobalQuality") ?? "auto",
                    DefaultFormat = (string)root.Attribute("GlobalFormat") ?? "png"
                };
            }
            catch
            {
                return defaults;
            }
        }

        public static void SaveSettings(List<AiImageApiSettings> settingsList, AiImageGlobalSettings globalSettings)
        {
            try
            {
                var root = new XElement(SettingsRootElement);

                if (globalSettings != null)
                {
                    root.SetAttributeValue("OverridePerModel", globalSettings.OverridePerModel);
                    root.SetAttributeValue("ModelPreviewCount", globalSettings.ModelPreviewCount);
                    root.SetAttributeValue("GlobalWidth", globalSettings.DefaultWidth);
                    root.SetAttributeValue("GlobalHeight", globalSettings.DefaultHeight);
                    root.SetAttributeValue("GlobalQuality", globalSettings.DefaultQuality ?? "auto");
                    root.SetAttributeValue("GlobalFormat", globalSettings.DefaultFormat ?? "png");
                }

                foreach (var s in settingsList ?? Enumerable.Empty<AiImageApiSettings>())
                {
                    root.Add(new XElement("AiImageApi",
                        new XElement("DisplayName", s.DisplayName ?? string.Empty),
                        new XElement("EndpointUrl", s.EndpointUrl ?? string.Empty),
                        new XElement("ApiToken", s.ApiToken ?? string.Empty),
                        new XElement("Model", s.Model ?? string.Empty),
                        new XElement("DefaultWidth", s.DefaultWidth),
                        new XElement("DefaultHeight", s.DefaultHeight),
                        new XElement("DefaultQuality", s.DefaultQuality ?? string.Empty),
                        new XElement("DefaultFormat", s.DefaultFormat ?? string.Empty),
                        new XElement("IconPath", s.IconPath ?? string.Empty),
                        new XElement("LockAspectRatio", s.LockAspectRatio)));
                }

                var path = GetSettingsFilePath();
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir))
                    Directory.CreateDirectory(dir);

                var doc = new XDocument(root);
                doc.Save(path);
            }
            catch
            {
            }
        }

        public static void SaveSettingsBounds(Rectangle? settingsBounds, Rectangle? globalSettingsBounds)
        {
            try
            {
                var path = GetSettingsFilePath();
                if (!File.Exists(path))
                    return;

                var doc = XDocument.Load(path);
                var root = doc.Root;
                if (root == null)
                    return;

                if (settingsBounds.HasValue)
                {
                    var b = settingsBounds.Value;
                    root.SetAttributeValue("AiSettingsEditorX", b.X);
                    root.SetAttributeValue("AiSettingsEditorY", b.Y);
                    root.SetAttributeValue("AiSettingsEditorWidth", b.Width);
                    root.SetAttributeValue("AiSettingsEditorHeight", b.Height);
                }

                if (globalSettingsBounds.HasValue)
                {
                    var b = globalSettingsBounds.Value;
                    root.SetAttributeValue("GlobalSettingsEditorX", b.X);
                    root.SetAttributeValue("GlobalSettingsEditorY", b.Y);
                    root.SetAttributeValue("GlobalSettingsEditorWidth", b.Width);
                    root.SetAttributeValue("GlobalSettingsEditorHeight", b.Height);
                }

                doc.Save(path);
            }
            catch
            {
            }
        }

        public static bool TryParseSettingsBounds(out Rectangle settingsBounds, out Rectangle globalSettingsBounds)
        {
            settingsBounds = Rectangle.Empty;
            globalSettingsBounds = Rectangle.Empty;
            bool hasSettings = false, hasGlobal = false;

            try
            {
                var path = GetSettingsFilePath();
                if (!File.Exists(path))
                    return false;

                var doc = XDocument.Load(path);
                var root = doc.Root;
                if (root == null)
                    return false;

                hasSettings = TryParseRect(root, "AiSettingsEditor", 500, 400, out settingsBounds);
                hasGlobal = TryParseRect(root, "GlobalSettingsEditor", 500, 350, out globalSettingsBounds);
            }
            catch
            {
            }

            return hasSettings || hasGlobal;
        }

        private static bool TryParseRect(XElement root, string prefix, int minW, int minH, out Rectangle rect)
        {
            rect = Rectangle.Empty;
            if (int.TryParse((string)root.Attribute(prefix + "X"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int x) &&
                int.TryParse((string)root.Attribute(prefix + "Y"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int y) &&
                int.TryParse((string)root.Attribute(prefix + "Width"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int w) &&
                int.TryParse((string)root.Attribute(prefix + "Height"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int h) &&
                w >= minW && h >= minH)
            {
                rect = new Rectangle(x, y, w, h);
                return true;
            }
            return false;
        }

        public static int ClampModelPreviewCount(int count)
        {
            return Math.Max(1, Math.Min(AiModelButtonCount, count));
        }

        public static AiImageApiSettings FindSettingsByModel(List<AiImageApiSettings> list, string model)
        {
            return list?.FirstOrDefault(x =>
                string.Equals(x.Model, model, StringComparison.OrdinalIgnoreCase));
        }

        public static bool TryGenerateImage(AiImageApiSettings settings, string prompt, int width, int height,
            string quality, string format, out string outputFilePath, out string errorMessage)
        {
            outputFilePath = null;
            errorMessage = string.Empty;

            try
            {
                if (settings == null || string.IsNullOrWhiteSpace(settings.EndpointUrl))
                {
                    errorMessage = "API 设置无效。";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(settings.ApiToken))
                {
                    errorMessage = "请先设置 API 令牌（按住 Ctrl 点击按钮打开设置）。";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(prompt))
                {
                    errorMessage = "请输入提示词。";
                    return false;
                }

                var requestBody = new StringBuilder();
                requestBody.Append("{");
                requestBody.Append("\"model\":\"" + EscapeJson(settings.Model) + "\",");
                requestBody.Append("\"prompt\":\"" + EscapeJson(prompt) + "\",");
                requestBody.Append("\"size\":\"" + width + "x" + height + "\",");
                requestBody.Append("\"quality\":\"" + EscapeJson(quality ?? settings.DefaultQuality ?? "auto") + "\",");
                requestBody.Append("\"output_format\":\"" + EscapeJson(format ?? settings.DefaultFormat ?? "png") + "\",");
                requestBody.Append("\"response_format\":\"b64_json\",");
                requestBody.Append("\"n\":1");
                requestBody.Append("}");

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var httpRequest = (HttpWebRequest)WebRequest.Create(settings.EndpointUrl);
                httpRequest.Method = "POST";
                httpRequest.ContentType = "application/json";
                httpRequest.Headers["Authorization"] = "Bearer " + settings.ApiToken;
                httpRequest.Timeout = 120000;

                var bodyBytes = Encoding.UTF8.GetBytes(requestBody.ToString());
                using (var requestStream = httpRequest.GetRequestStream())
                {
                    requestStream.Write(bodyBytes, 0, bodyBytes.Length);
                }

                using (var response = (HttpWebResponse)httpRequest.GetResponse())
                using (var responseStream = response.GetResponseStream())
                using (var reader = new StreamReader(responseStream, Encoding.UTF8))
                {
                    var responseJson = reader.ReadToEnd();

                    if (string.IsNullOrWhiteSpace(responseJson))
                    {
                        errorMessage = "API 返回了空响应。";
                        return false;
                    }

                    var base64Data = ExtractJsonStringValue(responseJson, "b64_json");
                    if (!string.IsNullOrWhiteSpace(base64Data))
                    {
                        return SaveBase64Image(base64Data, format, out outputFilePath, out errorMessage);
                    }

                    var imageUrl = ExtractJsonStringValue(responseJson, "url");
                    if (!string.IsNullOrWhiteSpace(imageUrl))
                    {
                        return DownloadImageFromUrl(imageUrl, format, settings.ApiToken, out outputFilePath, out errorMessage);
                    }

                    errorMessage = "API 响应中未找到图片数据。";
                    return false;
                }
            }
            catch (WebException ex)
            {
                try
                {
                    using (var errorStream = ex.Response?.GetResponseStream())
                    using (var errorReader = new StreamReader(errorStream ?? Stream.Null, Encoding.UTF8))
                    {
                        var errorBody = errorReader.ReadToEnd();
                        errorMessage = string.IsNullOrWhiteSpace(errorBody)
                            ? "API 请求失败：" + ex.Message
                            : "API 请求失败：" + errorBody;
                    }
                }
                catch
                {
                    errorMessage = "API 请求失败：" + ex.Message;
                }
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "生图失败：" + ex.Message;
                return false;
            }
        }

        private static bool DownloadImageFromUrl(string url, string format, string apiToken, out string filePath, out string error)
        {
            filePath = null;
            error = string.Empty;

            try
            {
                var ext = "." + (format ?? "png").TrimStart('.');
                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw", "AiImages");
                Directory.CreateDirectory(tempDir);
                filePath = Path.Combine(tempDir, "ai_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ext);

                using (var client = new WebClient())
                {
                    client.Headers["Authorization"] = "Bearer " + (apiToken ?? string.Empty);
                    client.DownloadFile(url, filePath);
                }

                if (File.Exists(filePath) && new FileInfo(filePath).Length > 0)
                    return true;

                error = "图片下载失败。";
                return false;
            }
            catch (Exception ex)
            {
                error = "图片下载失败：" + ex.Message;
                return false;
            }
        }

        private static bool SaveBase64Image(string base64, string format, out string filePath, out string error)
        {
            filePath = null;
            error = string.Empty;

            try
            {
                var bytes = Convert.FromBase64String(base64);
                var ext = "." + (format ?? "png").TrimStart('.');
                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw", "AiImages");
                Directory.CreateDirectory(tempDir);
                filePath = Path.Combine(tempDir, "ai_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ext);
                File.WriteAllBytes(filePath, bytes);

                if (File.Exists(filePath) && new FileInfo(filePath).Length > 0)
                    return true;

                error = "Base64 图片保存失败。";
                return false;
            }
            catch (Exception ex)
            {
                error = "Base64 图片保存失败：" + ex.Message;
                return false;
            }
        }

        private static void ParseLegacySize(string sizeStr, out int width, out int height)
        {
            width = 1024;
            height = 1024;
            if (string.IsNullOrWhiteSpace(sizeStr))
                return;
            var parts = sizeStr.Split('x');
            if (parts.Length == 2)
            {
                int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out width);
                int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out height);
            }
            if (width < 1) width = 1024;
            if (height < 1) height = 1024;
        }

        private static int ParseInt(string s, int defaultValue)
        {
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
                return result;
            return defaultValue;
        }

        private static bool ParseBool(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return false;
            return string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) || s == "1";
        }

        private static string ExtractJsonStringValue(string json, string key)
        {
            var search = "\"" + key + "\"";
            var keyIndex = json.IndexOf(search, StringComparison.Ordinal);
            if (keyIndex < 0)
                return null;

            var colonIndex = json.IndexOf(':', keyIndex + search.Length);
            if (colonIndex < 0)
                return null;

            var valueStart = json.IndexOf('"', colonIndex + 1);
            if (valueStart < 0)
                return null;

            var sb = new StringBuilder();
            for (int i = valueStart + 1; i < json.Length; i++)
            {
                var ch = json[i];
                if (ch == '\\' && i + 1 < json.Length)
                {
                    i++;
                    var next = json[i];
                    switch (next)
                    {
                        case '"': sb.Append('"'); break;
                        case '\\': sb.Append('\\'); break;
                        case '/': sb.Append('/'); break;
                        case 'n': sb.Append('\n'); break;
                        case 'r': sb.Append('\r'); break;
                        case 't': sb.Append('\t'); break;
                        case 'u':
                            if (i + 4 < json.Length)
                            {
                                var hex = json.Substring(i + 1, 4);
                                if (int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code))
                                {
                                    sb.Append((char)code);
                                    i += 4;
                                }
                            }
                            break;
                        default: sb.Append(next); break;
                    }
                }
                else if (ch == '"')
                {
                    break;
                }
                else
                {
                    sb.Append(ch);
                }
            }

            return sb.ToString();
        }

        private static string EscapeJson(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t");
        }

        private static string GetSettingsFilePath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BioDraw",
                "AiImageSettings.xml");
        }
    }
}
