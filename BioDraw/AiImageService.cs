using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Xml.Linq;

namespace BioDraw
{
    internal static class AiImageService
    {
        private const string SettingsRootElement = "BioDrawAiImageSettings";
        public const int AiModelButtonCount = 12;
        public const int DefaultModelPreviewCount = 5;


        private static readonly Lazy<HttpClient> SharedHttpClient = new Lazy<HttpClient>(() =>
        {
            // ApiMart requires TLS 1.2+; .NET Framework 4.7.2 does not enable it by default.
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var client = new HttpClient();
            client.Timeout = TimeSpan.FromMinutes(5);
            return client;
        });

        public static AiImageApiSettings CreateDefaultSettings()
        {
            return new AiImageApiSettings
            {
                DisplayName = "GPT-Image-2",
                EndpointUrl = "https://api.apimart.ai/v1/images/generations",
                ApiToken = string.Empty,
                Model = "gpt-image-2",
                DefaultWidth = 1024,
                DefaultHeight = 1024,
                DefaultQuality = "auto",
                DefaultFormat = "png",
                IconPath = string.Empty,
                LockAspectRatio = false,
                Resolution = "2k"
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
                DefaultFormat = "png",
                ImgOverridePerModel = false,
                ImgModelPreviewCount = DefaultModelPreviewCount,
                ImgDefaultWidth = 1024,
                ImgDefaultHeight = 1024,
                ImgDefaultQuality = "auto",
                ImgDefaultFormat = "png"
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
                        EndpointUrl = (string)entry.Element("EndpointUrl") ?? "https://api.apimart.ai/v1/images/generations",
                        ApiToken = (string)entry.Element("ApiToken") ?? string.Empty,
                        Model = (string)entry.Element("Model") ?? "gpt-image-2",
                        DefaultQuality = (string)entry.Element("DefaultQuality") ?? "auto",
                        DefaultFormat = (string)entry.Element("DefaultFormat") ?? "png",
                        IconPath = (string)entry.Element("IconPath") ?? string.Empty,
                        LockAspectRatio = ParseBool((string)entry.Element("LockAspectRatio")),
                        Resolution = (string)entry.Element("Resolution") ?? "2k"
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
                    DefaultFormat = (string)root.Attribute("GlobalFormat") ?? "png",
                    ImgOverridePerModel = ParseBool((string)root.Attribute("ImgOverridePerModel")),
                    ImgModelPreviewCount = ClampModelPreviewCount(ParseInt((string)root.Attribute("ImgModelPreviewCount"), DefaultModelPreviewCount)),
                    ImgDefaultWidth = ParseInt((string)root.Attribute("ImgDefaultWidth"), 1024),
                    ImgDefaultHeight = ParseInt((string)root.Attribute("ImgDefaultHeight"), 1024),
                    ImgDefaultQuality = (string)root.Attribute("ImgDefaultQuality") ?? "auto",
                    ImgDefaultFormat = (string)root.Attribute("ImgDefaultFormat") ?? "png"
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
                    root.SetAttributeValue("ImgOverridePerModel", globalSettings.ImgOverridePerModel);
                    root.SetAttributeValue("ImgModelPreviewCount", globalSettings.ImgModelPreviewCount);
                    root.SetAttributeValue("ImgDefaultWidth", globalSettings.ImgDefaultWidth);
                    root.SetAttributeValue("ImgDefaultHeight", globalSettings.ImgDefaultHeight);
                    root.SetAttributeValue("ImgDefaultQuality", globalSettings.ImgDefaultQuality ?? "auto");
                    root.SetAttributeValue("ImgDefaultFormat", globalSettings.ImgDefaultFormat ?? "png");
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
                        new XElement("LockAspectRatio", s.LockAspectRatio),
                        new XElement("Resolution", s.Resolution ?? "2k")));
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


        public static void SaveSettingsBounds(Rectangle? settingsBounds, Rectangle? globalSettingsBounds,
            Rectangle? imgGlobalSettingsBounds = null)
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

                if (imgGlobalSettingsBounds.HasValue)
                {
                    var b = imgGlobalSettingsBounds.Value;
                    root.SetAttributeValue("ImgGlobalSettingsEditorX", b.X);
                    root.SetAttributeValue("ImgGlobalSettingsEditorY", b.Y);
                    root.SetAttributeValue("ImgGlobalSettingsEditorWidth", b.Width);
                    root.SetAttributeValue("ImgGlobalSettingsEditorHeight", b.Height);
                }

                doc.Save(path);
            }
            catch
            {
            }
        }

        public static bool TryParseSettingsBounds(out Rectangle settingsBounds, out Rectangle globalSettingsBounds,
            out Rectangle imgGlobalSettingsBounds)
        {
            settingsBounds = Rectangle.Empty;
            globalSettingsBounds = Rectangle.Empty;
            imgGlobalSettingsBounds = Rectangle.Empty;

            try
            {
                var path = GetSettingsFilePath();
                if (!File.Exists(path))
                    return false;

                var doc = XDocument.Load(path);
                var root = doc.Root;
                if (root == null)
                    return false;

                TryParseRect(root, "AiSettingsEditor", 500, 400, out settingsBounds);
                TryParseRect(root, "GlobalSettingsEditor", 500, 350, out globalSettingsBounds);
                TryParseRect(root, "ImgGlobalSettingsEditor", 500, 350, out imgGlobalSettingsBounds);
            }
            catch
            {
            }

            return true;
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


        // ==============================================================
        // ApiMart async task polling infrastructure
        // ==============================================================
        private const string ApiMartBaseUrl = "https://api.apimart.ai";
        private const int TaskPollTimeoutMs = 360000; // 6 min
        private const int InitialTaskPollDelayMs = 12000; // ApiMart recommends first query after 10-20 sec
        private const int TaskPollIntervalMs = 3000;   // 3 sec

        private static bool IsApiMartEndpoint(string url)
        {
            return url != null && url.IndexOf("apimart.ai", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// GPT-Image-2 uses aspect ratios in the "size" field rather than pixel dimensions.
        /// </summary>
        private static bool ModelUsesAspectRatio(string model)
        {
            return string.Equals(model, "gpt-image-2", StringComparison.OrdinalIgnoreCase)
                || string.Equals(model, "gpt-image-2-official", StringComparison.OrdinalIgnoreCase)
                || string.Equals(model, "gpt-image-1", StringComparison.OrdinalIgnoreCase);
        }

        private static readonly string[] GptImage2Ratios =
        {
            "1:1", "4:3", "3:4", "3:2", "2:3", "5:4", "4:5", "16:9", "9:16", "21:9"
        };

        private static string ConvertToAspectRatio(int width, int height)
        {
            if (width <= 0 || height <= 0)
                return "1:1";

            double target = (double)width / height;
            string best = "1:1";
            double bestDiff = double.MaxValue;

            foreach (string ratio in GptImage2Ratios)
            {
                string[] parts = ratio.Split(':');
                if (parts.Length != 2) continue;
                if (!double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double w) ||
                    !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double h) ||
                    h == 0)
                    continue;

                double r = w / h;
                double diff = Math.Abs(r - target);
                if (diff < bestDiff)
                {
                    bestDiff = diff;
                    best = ratio;
                }
            }

            return best;
        }

        /// <summary>
        /// Builds the JSON request body for ApiMart image generation.
        /// For GPT-Image-2: "size" is aspect ratio, "resolution" controls output pixels.
        /// For other models: "size" is "WxH" pixel dimensions, no resolution field.
        /// When omitSize is true (img-to-img with lock ratio), size is omitted so the API
        /// preserves the reference image's aspect ratio at the chosen resolution bucket.
        /// </summary>
        private static string BuildImageRequestBody(
            string model, string prompt, int width, int height,
            string quality, string format, string resolution,
            int n, string[] imageUrls, bool useGptImage2Style, bool omitSize = false)
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append("\"model\":\"" + EscapeJson(model ?? "gpt-image-2") + "\",");
            sb.Append("\"prompt\":\"" + EscapeJson(prompt ?? string.Empty) + "\",");

            if (omitSize && imageUrls != null && imageUrls.Length > 0)
            {
                if (useGptImage2Style)
                {
                    string res = resolution ?? "2k";
                    if (res != "1k" && res != "2k" && res != "4k")
                        res = "2k";
                    sb.Append("\"resolution\":\"" + res + "\",");
                }
            }
            else if (useGptImage2Style)
            {
                string aspectRatio = ConvertToAspectRatio(width, height);
                sb.Append("\"size\":\"" + aspectRatio + "\",");

                string res = resolution ?? "2k";
                if (res != "1k" && res != "2k" && res != "4k")
                    res = "2k";
                sb.Append("\"resolution\":\"" + res + "\",");
            }
            else
            {
                sb.Append("\"size\":\"" + width + "x" + height + "\",");
            }

            if (!string.IsNullOrWhiteSpace(quality) && !string.Equals(quality, "auto", StringComparison.OrdinalIgnoreCase))
                sb.Append("\"quality\":\"" + EscapeJson(quality) + "\",");

            sb.Append("\"output_format\":\"" + EscapeJson(format ?? "png") + "\",");

            if (imageUrls != null && imageUrls.Length > 0)
            {
                sb.Append("\"image_urls\":[");
                for (int i = 0; i < imageUrls.Length; i++)
                {
                    if (i > 0) sb.Append(",");
                    sb.Append("\"" + EscapeJson(imageUrls[i]) + "\"");
                }
                sb.Append("],");
            }

            sb.Append("\"n\":" + Math.Max(1, Math.Min(4, n)) + "");
            sb.Append("}");

            return sb.ToString();
        }

        private static string SubmitApiMartTask(
            string endpointUrl, string apiToken, string requestJson,
            out string errorMessage)
        {
            errorMessage = string.Empty;

            try
            {
                var bodyBytes = Encoding.UTF8.GetBytes(requestJson);
                using (var content = new ByteArrayContent(bodyBytes))
                {
                    content.Headers.ContentType =
                        new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");

                    using (var request = new HttpRequestMessage(HttpMethod.Post, endpointUrl))
                    {
                        request.Headers.Authorization =
                            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiToken);
                        request.Content = content;

                        using (var response = SharedHttpClient.Value.SendAsync(request).Result)
                        {
                            var responseJson = response.Content.ReadAsStringAsync().Result;

                            if (!response.IsSuccessStatusCode)
                            {
                                errorMessage = "Task submission " + FormatApiError((int)response.StatusCode, responseJson, requestJson);
                                return null;
                            }

                            if (string.IsNullOrWhiteSpace(responseJson))
                            {
                                errorMessage = "API returned empty response.";
                                return null;
                            }

                            string taskId = ExtractApiMartTaskId(responseJson);
                            if (string.IsNullOrWhiteSpace(taskId))
                            {
                                errorMessage = "Cannot extract task_id from API response: " + responseJson;
                                return null;
                            }

                            return taskId;
                        }
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                errorMessage = "API request failed (network): " + ex.Message;
                return null;
            }
            catch (TaskCanceledException)
            {
                errorMessage = "API request timed out.";
                return null;
            }
            catch (AggregateException ae)
            {
                var inner = ae.InnerException ?? ae;
                if (inner is TaskCanceledException)
                    errorMessage = "API request timed out: " + inner.Message;
                else
                    errorMessage = "API request failed: " + inner.Message;
                return null;
            }
            catch (Exception ex)
            {
                errorMessage = "Task submission failed: " + ex.Message;
                return null;
            }
        }

        private static bool PollApiMartTaskForImages(
            string apiToken, string taskId, string format,
            out string outputFilePath, out string errorMessage)
        {
            outputFilePath = null;
            errorMessage = string.Empty;

            var pollUrl = ApiMartBaseUrl + "/v1/tasks/" + taskId;
            var deadline = Environment.TickCount + TaskPollTimeoutMs;
            string lastResponseJson = string.Empty;

            try
            {
                if (InitialTaskPollDelayMs > 0)
                    System.Threading.Thread.Sleep(InitialTaskPollDelayMs);

                while (Environment.TickCount < deadline)
                {
                    using (var request = new HttpRequestMessage(HttpMethod.Get, pollUrl))
                    {
                        request.Headers.Authorization =
                            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiToken ?? string.Empty);

                        using (var response = SharedHttpClient.Value.SendAsync(request).Result)
                        {
                            var responseJson = response.Content.ReadAsStringAsync().Result;
                            if (string.IsNullOrWhiteSpace(responseJson))
                            {
                                errorMessage = "Task polling returned empty response.";
                                return false;
                            }

                            lastResponseJson = responseJson;

                            if (!response.IsSuccessStatusCode)
                            {
                                errorMessage = "Task polling " + FormatApiError((int)response.StatusCode, responseJson, null);
                                return false;
                            }

                            string status = ExtractTaskStatus(responseJson);

                            string[] imageUrls = ExtractImageUrlsFromTaskResult(responseJson);
                            if (imageUrls != null && imageUrls.Length > 0 && !IsFailedTaskStatus(status))
                                return DownloadTaskResultImage(lastResponseJson, apiToken, format, out outputFilePath, out errorMessage);

                            if (IsCompletedTaskStatus(status))
                                return DownloadTaskResultImage(lastResponseJson, apiToken, format, out outputFilePath, out errorMessage);

                            if (IsFailedTaskStatus(status))
                            {
                                errorMessage = "Image generation task failed. Response: " + responseJson;
                                return false;
                            }
                        }
                    }

                    System.Threading.Thread.Sleep(TaskPollIntervalMs);
                }

                errorMessage = "Task polling timed out (" + (TaskPollTimeoutMs / 60000) + " min). Last status: " + lastResponseJson;
                return false;
            }
            catch (TaskCanceledException)
            {
                errorMessage = "Task polling timed out.";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "Task polling failed: " + ex.Message;
                return false;
            }
        }

        private static bool DownloadTaskResultImage(
            string taskResultJson, string apiToken, string format,
            out string outputFilePath, out string errorMessage)
        {
            outputFilePath = null;
            errorMessage = string.Empty;

            try
            {
                string[] imageUrls = ExtractImageUrlsFromTaskResult(taskResultJson);
                if (imageUrls == null || imageUrls.Length == 0)
                {
                    errorMessage = "Task completed but no image URLs found. Response: " + taskResultJson;
                    return false;
                }

                string imageUrl = imageUrls[0];

                var ext = "." + (format ?? "png").TrimStart('.');
                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw", "AiImages");
                Directory.CreateDirectory(tempDir);
                outputFilePath = Path.Combine(tempDir,
                    "ai_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" +
                    Guid.NewGuid().ToString("N").Substring(0, 6) + ext);

                using (var request = new HttpRequestMessage(HttpMethod.Get, imageUrl))
                {
                    if (!string.IsNullOrWhiteSpace(apiToken))
                        request.Headers.Authorization =
                            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiToken);

                    using (var response = SharedHttpClient.Value.SendAsync(request).Result)
                    {
                        response.EnsureSuccessStatusCode();
                        var bytes = response.Content.ReadAsByteArrayAsync().Result;
                        File.WriteAllBytes(outputFilePath, bytes);
                    }
                }

                if (File.Exists(outputFilePath) && new FileInfo(outputFilePath).Length > 0)
                    return true;

                errorMessage = "Downloaded image file is empty.";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "Failed to download task result image: " + ex.Message;
                return false;
            }
        }

        private static string[] ExtractImageUrlsFromTaskResult(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return null;

            var parsed = TryParseJsonObject(json);
            var parsedUrls = new List<string>();
            if (CollectHttpUrls(parsed, parsedUrls) && parsedUrls.Count > 0)
                return parsedUrls.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

            var urls = new List<string>();
            int searchStart = 0;
            while (true)
            {
                int urlKeyIdx = json.IndexOf("\"url\"", searchStart, StringComparison.Ordinal);
                if (urlKeyIdx < 0) break;
                int colonIdx = json.IndexOf(':', urlKeyIdx + 5);
                if (colonIdx < 0) break;
                int bracketIdx = json.IndexOf('[', colonIdx + 1);
                if (bracketIdx < 0) { searchStart = urlKeyIdx + 5; continue; }
                int closeBracketIdx = json.IndexOf(']', bracketIdx + 1);
                if (closeBracketIdx < 0) break;

                string arrayContent = json.Substring(bracketIdx + 1, closeBracketIdx - bracketIdx - 1).Trim();
                if (!string.IsNullOrWhiteSpace(arrayContent))
                {
                    string[] items = SplitJsonArray(arrayContent);
                    foreach (string item in items)
                    {
                        string trimmed = item.Trim().Trim('"');
                        if (!string.IsNullOrWhiteSpace(trimmed) && trimmed.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                            urls.Add(trimmed);
                    }
                }
                searchStart = closeBracketIdx + 1;
            }

            return urls.Count > 0 ? urls.ToArray() : null;
        }

        private static string ExtractTaskStatus(string json)
        {
            var parsed = TryParseJsonObject(json);
            var status = FindFirstStringByKey(parsed, "status", "task_status", "state");
            if (!string.IsNullOrWhiteSpace(status))
                return status;

            status = ExtractJsonStringValue(json, "status");
            if (!string.IsNullOrWhiteSpace(status))
                return status;

            return ExtractJsonStringValueNested(json, "status");
        }

        private static string ExtractJsonStringValueNested(string json, string key)
        {
            var val = ExtractJsonStringValue(json, key);
            if (!string.IsNullOrWhiteSpace(val)) return val;

            int dataIdx = json.IndexOf("\"data\"", StringComparison.Ordinal);
            if (dataIdx < 0) return null;
            int colonIdx = json.IndexOf(':', dataIdx + 6);
            if (colonIdx < 0) return null;
            int braceIdx = json.IndexOf('{', colonIdx + 1);
            if (braceIdx < 0) return null;

            int depth = 1;
            int closeBrace = braceIdx + 1;
            for (; closeBrace < json.Length; closeBrace++)
            {
                if (json[closeBrace] == '{') depth++;
                else if (json[closeBrace] == '}') depth--;
                if (depth == 0) break;
            }
            if (closeBrace >= json.Length) return null;
            string nestedJson = json.Substring(braceIdx, closeBrace - braceIdx + 1);

            return ExtractJsonStringValue(nestedJson, key);
        }

        private static string ExtractApiMartTaskId(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return null;

            var parsed = TryParseJsonObject(json);
            string parsedTaskId = FindFirstStringByKey(parsed, "task_id", "id");
            if (!string.IsNullOrWhiteSpace(parsedTaskId))
                return parsedTaskId;

            // 1) "id" field -- most common (OpenAI-compatible, ApiMart flat format)
            string id = ExtractJsonStringValue(json, "id");
            if (!string.IsNullOrWhiteSpace(id))
                return id;

            // 2) "task_id" field -- alternate ApiMart format
            string taskId = ExtractJsonStringValue(json, "task_id");
            if (!string.IsNullOrWhiteSpace(taskId))
                return taskId;

            // 3) "data.id" nested form (some endpoints wrap in "data")
            string dataId = ExtractJsonStringValueNested(json, "id");
            if (!string.IsNullOrWhiteSpace(dataId))
                return dataId;

            // 4) "data.task_id" nested form
            string dataTaskId = ExtractJsonStringValueNested(json, "task_id");
            if (!string.IsNullOrWhiteSpace(dataTaskId))
                return dataTaskId;

            // 5) Fallback: search for "task_" prefix in string VALUES only.
            //    Avoid matching key names like "task_id" by checking that the
            //    character just before "task_" is a double-quote (value start).
            int searchPos = 0;
            while (searchPos < json.Length)
            {
                int taskIdx = json.IndexOf("task_", searchPos, StringComparison.Ordinal);
                if (taskIdx < 0)
                    break;

                // Only match when preceded by '"' (value, not key name)
                if (taskIdx > 0 && json[taskIdx - 1] == '"')
                {
                    int end = taskIdx + 5;
                    while (end < json.Length)
                    {
                        char c = json[end];
                        if (c == '"' || c == ',' || c == '}' || c == ']' || c == '\n' || c == '\r')
                            break;
                        end++;
                    }
                    if (end > taskIdx)
                        return json.Substring(taskIdx, end - taskIdx);
                }

                searchPos = taskIdx + 5;
            }

            return null;
        }

        private static bool IsCompletedTaskStatus(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
                return false;

            return string.Equals(status, "completed", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "succeeded", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "success", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "done", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "finished", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsFailedTaskStatus(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
                return false;

            return string.Equals(status, "failed", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "error", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "cancelled", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "canceled", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "expired", StringComparison.OrdinalIgnoreCase)
                || string.Equals(status, "rejected", StringComparison.OrdinalIgnoreCase);
        }

        private static object TryParseJsonObject(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return null;

            try
            {
                return new JavaScriptSerializer().DeserializeObject(json);
            }
            catch
            {
                return null;
            }
        }

        private static string FindFirstStringByKey(object node, params string[] keys)
        {
            if (node == null || keys == null || keys.Length == 0)
                return null;

            var visited = new HashSet<object>();
            return FindFirstStringByKeyCore(node, keys, visited);
        }

        private static string FindFirstStringByKeyCore(object node, string[] keys, HashSet<object> visited)
        {
            if (node == null)
                return null;

            if (node is string)
                return null;

            if (!visited.Add(node))
                return null;

            var dict = node as Dictionary<string, object>;
            if (dict != null)
            {
                foreach (string key in keys)
                {
                    if (TryGetDictionaryValue(dict, key, out object value))
                    {
                        string stringValue = ExtractFirstStringValue(value);
                        if (!string.IsNullOrWhiteSpace(stringValue))
                            return stringValue;
                    }
                }

                foreach (var value in dict.Values)
                {
                    string nested = FindFirstStringByKeyCore(value, keys, visited);
                    if (!string.IsNullOrWhiteSpace(nested))
                        return nested;
                }

                return null;
            }

            var array = node as object[];
            if (array != null)
            {
                foreach (var item in array)
                {
                    string nested = FindFirstStringByKeyCore(item, keys, visited);
                    if (!string.IsNullOrWhiteSpace(nested))
                        return nested;
                }
            }

            return null;
        }

        private static bool CollectHttpUrls(object node, List<string> urls)
        {
            if (node == null || urls == null)
                return false;

            bool found = false;
            var dict = node as Dictionary<string, object>;
            if (dict != null)
            {
                foreach (var kvp in dict)
                    found |= CollectHttpUrls(kvp.Value, urls);
                return found;
            }

            var array = node as object[];
            if (array != null)
            {
                foreach (var item in array)
                    found |= CollectHttpUrls(item, urls);
                return found;
            }

            var str = node as string;
            if (!string.IsNullOrWhiteSpace(str) && str.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                urls.Add(str);
                return true;
            }

            return false;
        }

        private static bool TryGetDictionaryValue(Dictionary<string, object> dict, string key, out object value)
        {
            value = null;
            if (dict == null || string.IsNullOrWhiteSpace(key))
                return false;

            foreach (var kvp in dict)
            {
                if (string.Equals(kvp.Key, key, StringComparison.OrdinalIgnoreCase))
                {
                    value = kvp.Value;
                    return true;
                }
            }

            return false;
        }

        private static string ExtractFirstStringValue(object value)
        {
            if (value == null)
                return null;

            var str = value as string;
            if (!string.IsNullOrWhiteSpace(str))
                return str;

            var array = value as object[];
            if (array != null)
            {
                foreach (var item in array)
                {
                    string nested = ExtractFirstStringValue(item);
                    if (!string.IsNullOrWhiteSpace(nested))
                        return nested;
                }
            }

            var dict = value as Dictionary<string, object>;
            if (dict != null)
            {
                foreach (var nestedValue in dict.Values)
                {
                    string nested = ExtractFirstStringValue(nestedValue);
                    if (!string.IsNullOrWhiteSpace(nested))
                        return nested;
                }
            }

            return null;
        }

        private static string[] SplitJsonArray(string arrayContent)
        {
            if (string.IsNullOrWhiteSpace(arrayContent))
                return new string[0];

            var items = new List<string>();
            int depth = 0;
            bool inString = false;
            int start = 0;

            for (int i = 0; i < arrayContent.Length; i++)
            {
                char c = arrayContent[i];
                if (inString)
                {
                    if (c == '\\' && i + 1 < arrayContent.Length) { i++; continue; }
                    if (c == '"') inString = false;
                    continue;
                }
                if (c == '"') inString = true;
                else if (c == '{' || c == '[') depth++;
                else if (c == '}' || c == ']') depth--;
                else if (c == ',' && depth == 0)
                {
                    items.Add(arrayContent.Substring(start, i - start));
                    start = i + 1;
                }
            }
            if (start < arrayContent.Length)
                items.Add(arrayContent.Substring(start));

            return items.ToArray();
        }

        private static string FormatApiError(int statusCode, string responseJson, string requestJson)
        {
            var sb = new StringBuilder();
            sb.Append("API request failed (" + statusCode + ")");
            if (!string.IsNullOrWhiteSpace(responseJson))
            {
                var errMsg = ExtractJsonStringValue(responseJson, "message")
                          ?? ExtractJsonStringValue(responseJson, "error")
                          ?? ExtractJsonStringValue(responseJson, "errMsg");
                if (!string.IsNullOrWhiteSpace(errMsg))
                    sb.Append(": " + errMsg);
                else if (responseJson.Length <= 500)
                    sb.Append(": " + responseJson);
                else
                    sb.Append(": " + responseJson.Substring(0, 500) + "...");
            }
            return sb.ToString();
        }


        // ==============================================================
        // Public API: TryGenerateImage
        // ==============================================================
        public static bool TryGenerateImage(AiImageApiSettings settings, string prompt, int width, int height,
            string quality, string format, out string outputFilePath, out string errorMessage)
        {
            outputFilePath = null;
            errorMessage = string.Empty;

            try
            {
                if (settings == null || string.IsNullOrWhiteSpace(settings.EndpointUrl))
                {
                    errorMessage = "API settings invalid.";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(settings.ApiToken))
                {
                    errorMessage = "Please set your API token first (Ctrl+click the button to open settings).";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(prompt))
                {
                    errorMessage = "Please enter a prompt.";
                    return false;
                }

                string model = settings.Model ?? "gpt-image-2";
                bool isApiMart = IsApiMartEndpoint(settings.EndpointUrl);
                bool useGpt2Style = ModelUsesAspectRatio(model);

                string resolution = settings.Resolution ?? "2k";
                string requestJson = BuildImageRequestBody(
                    model, prompt, width, height, quality, format, resolution,
                    n: 1, imageUrls: null, useGptImage2Style: useGpt2Style);

                if (isApiMart)
                {
                    string taskId = SubmitApiMartTask(
                        settings.EndpointUrl, settings.ApiToken, requestJson,
                        out errorMessage);

                    if (taskId == null)
                        return false;

                    return PollApiMartTaskForImages(
                        settings.ApiToken, taskId, format,
                        out outputFilePath, out errorMessage);
                }
                else
                {
                    return LegacySyncImageRequest(
                        settings.EndpointUrl, settings.ApiToken, requestJson,
                        format, out outputFilePath, out errorMessage);
                }
            }
            catch (HttpRequestException ex)
            {
                errorMessage = "API request failed (network): " + ex.Message;
                return false;
            }
            catch (TaskCanceledException ex)
            {
                errorMessage = "API request timed out: " + ex.Message;
                return false;
            }
            catch (AggregateException ae)
            {
                var inner = ae.InnerException ?? ae;
                if (inner is TaskCanceledException)
                    errorMessage = "API request timed out: " + inner.Message;
                else
                    errorMessage = "API request failed: " + inner.Message;
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "Image generation failed: " + ex.Message;
                return false;
            }
        }



        // ==============================================================
        // Public API: TryGenerateImageFromImage
        // ==============================================================
        public static bool TryGenerateImageFromImage(AiImageApiSettings settings, string prompt,
            int width, int height, string quality, string format, string sourceImagePath,
            bool lockAspectRatio, out string outputFilePath, out string errorMessage)
        {
            outputFilePath = null;
            errorMessage = string.Empty;

            try
            {
                if (settings == null || string.IsNullOrWhiteSpace(settings.EndpointUrl))
                {
                    errorMessage = "API settings invalid.";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(settings.ApiToken))
                {
                    errorMessage = "Please set your API token.";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(prompt))
                {
                    errorMessage = "Please enter a prompt.";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(sourceImagePath) || !File.Exists(sourceImagePath))
                {
                    errorMessage = "Source image not obtained.";
                    return false;
                }

                string model = settings.Model ?? "gpt-image-2";
                bool isApiMart = IsApiMartEndpoint(settings.EndpointUrl);
                bool useGpt2Style = ModelUsesAspectRatio(model);

                // Official ApiMart flow: upload the original image file to get a stable
                // URL, then reference it via image_urls. Base64 is no longer supported by
                // the generation API and inflates the request body, so we never inline it.
                string imageUrl = UploadImageToApiMart(
                    settings.EndpointUrl, settings.ApiToken, sourceImagePath, out errorMessage);
                if (string.IsNullOrWhiteSpace(imageUrl))
                    return false;

                var imageUrls = new[] { imageUrl };
                string resolution = settings.Resolution ?? "2k";
                string requestJson = BuildImageRequestBody(
                    model, prompt, width, height, quality, format, resolution,
                    n: 1, imageUrls: imageUrls, useGptImage2Style: useGpt2Style,
                    omitSize: lockAspectRatio);

                if (isApiMart)
                {
                    string taskId = SubmitApiMartTask(
                        settings.EndpointUrl, settings.ApiToken, requestJson,
                        out errorMessage);

                    if (taskId == null)
                        return false;

                    return PollApiMartTaskForImages(
                        settings.ApiToken, taskId, format,
                        out outputFilePath, out errorMessage);
                }
                else
                {
                    return LegacySyncImageRequest(
                        settings.EndpointUrl, settings.ApiToken, requestJson,
                        format, out outputFilePath, out errorMessage);
                }
            }
            catch (AggregateException ae)
            {
                var inner = ae.InnerException ?? ae;
                errorMessage = "API request failed: " + inner.Message;
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "Image-to-image failed: " + ex.Message;
                return false;
            }
        }



        // ==============================================================
        // Image upload (POST /v1/uploads/images) — official ApiMart flow
        // Uploads the original file as multipart/form-data (field "file"),
        // returns a stable public URL (valid ~72h) for use in image_urls.
        // No base64, no re-encoding: the bytes on disk are sent verbatim.
        // ==============================================================
        private static string UploadImageToApiMart(
            string generationEndpointUrl, string apiToken, string sourceImagePath,
            out string errorMessage)
        {
            errorMessage = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(sourceImagePath) || !File.Exists(sourceImagePath))
                {
                    errorMessage = "源图片文件不存在。";
                    return null;
                }

                // ApiMart only accepts JPEG/PNG/WebP/GIF. Convert unsupported formats to PNG.
                string uploadPath = EnsureUploadableFormat(sourceImagePath);

                string uploadUrl = DeriveUploadEndpoint(generationEndpointUrl);
                var fileBytes = File.ReadAllBytes(uploadPath);
                var fileName = Path.GetFileName(uploadPath);
                var mimeType = GetMimeTypeFromExtension(Path.GetExtension(uploadPath));

                using (var form = new MultipartFormDataContent())
                {
                    var fileContent = new ByteArrayContent(fileBytes);
                    fileContent.Headers.ContentType =
                        new System.Net.Http.Headers.MediaTypeHeaderValue(mimeType);
                    form.Add(fileContent, "file", fileName);

                    using (var request = new HttpRequestMessage(HttpMethod.Post, uploadUrl))
                    {
                        request.Headers.Authorization =
                            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiToken ?? string.Empty);
                        request.Content = form;

                        using (var response = SharedHttpClient.Value.SendAsync(request).Result)
                        {
                            var responseJson = response.Content.ReadAsStringAsync().Result;

                            if (!response.IsSuccessStatusCode)
                            {
                                errorMessage = "图片上传失败 " + FormatApiError((int)response.StatusCode, responseJson, null);
                                return null;
                            }

                            if (string.IsNullOrWhiteSpace(responseJson))
                            {
                                errorMessage = "图片上传返回空响应。";
                                return null;
                            }

                            string url = ExtractJsonStringValue(responseJson, "url");
                            if (string.IsNullOrWhiteSpace(url))
                            {
                                var parsed = TryParseJsonObject(responseJson);
                                var collected = new List<string>();
                                if (CollectHttpUrls(parsed, collected) && collected.Count > 0)
                                    url = collected[0];
                            }

                            if (string.IsNullOrWhiteSpace(url))
                            {
                                errorMessage = "无法从上传响应中提取图片 URL: " + responseJson;
                                return null;
                            }

                            return url;
                        }
                    }
                }
            }
            catch (AggregateException ae)
            {
                errorMessage = "图片上传失败：" + FlattenExceptionMessage(ae);
                return null;
            }
            catch (Exception ex)
            {
                errorMessage = "图片上传失败：" + FlattenExceptionMessage(ex);
                return null;
            }
        }

        /// <summary>
        /// If the source file is a format unsupported by ApiMart's upload endpoint
        /// (EMF/WMF/BMP/TIFF/etc.), converts it to PNG losslessly and returns the
        /// new path. For supported formats (JPEG/PNG/WebP/GIF) returns the original path.
        /// </summary>
        private static string EnsureUploadableFormat(string filePath)
        {
            var ext = (Path.GetExtension(filePath) ?? string.Empty).TrimStart('.').ToLowerInvariant();
            switch (ext)
            {
                case "jpg":
                case "jpeg":
                case "jfif":
                case "png":
                case "webp":
                case "gif":
                    return filePath;
            }

            // Unsupported format: convert to PNG via GDI+
            try
            {
                using (var img = Image.FromFile(filePath))
                {
                    var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw", "ImgToImg");
                    Directory.CreateDirectory(tempDir);
                    var pngPath = Path.Combine(tempDir,
                        Path.GetFileNameWithoutExtension(filePath) + "_" +
                        Guid.NewGuid().ToString("N").Substring(0, 6) + ".png");
                    img.Save(pngPath, ImageFormat.Png);
                    return pngPath;
                }
            }
            catch
            {
                return filePath;
            }
        }

        /// <summary>
        /// Flattens the exception chain into a single readable message for display.
        /// </summary>
        private static string FlattenExceptionMessage(Exception ex)
        {
            if (ex == null) return string.Empty;
            var sb = new StringBuilder();
            var current = ex is AggregateException ae ? (ae.InnerException ?? ex) : ex;
            while (current != null)
            {
                if (sb.Length > 0) sb.Append(" → ");
                sb.Append(current.Message);
                current = current.InnerException;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Derives the /v1/uploads/images endpoint from the configured generation
        /// endpoint, preserving the same host/scheme (e.g. proxies).
        /// </summary>
        private static string DeriveUploadEndpoint(string generationEndpointUrl)
        {
            const string uploadsPath = "/v1/uploads/images";

            if (!string.IsNullOrWhiteSpace(generationEndpointUrl))
            {
                int idx = generationEndpointUrl.IndexOf("/v1/images/generations", StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                    return generationEndpointUrl.Substring(0, idx) + uploadsPath;

                try
                {
                    var uri = new Uri(generationEndpointUrl, UriKind.Absolute);
                    return uri.GetLeftPart(UriPartial.Authority) + uploadsPath;
                }
                catch
                {
                }
            }

            return ApiMartBaseUrl + uploadsPath;
        }

        private static string GetMimeTypeFromExtension(string extension)
        {
            switch ((extension ?? string.Empty).TrimStart('.').ToLowerInvariant())
            {
                case "jpg":
                case "jpeg":
                case "jfif":
                    return "image/jpeg";
                case "webp":
                    return "image/webp";
                case "gif":
                    return "image/gif";
                case "bmp":
                    return "image/bmp";
                case "tif":
                case "tiff":
                    return "image/tiff";
                default:
                    return "image/png";
            }
        }


        // ==============================================================
        // Legacy synchronous flow (non-ApiMart endpoints)
        // ==============================================================
        private static bool LegacySyncImageRequest(
            string endpointUrl, string apiToken, string requestJson,
            string format, out string outputFilePath, out string errorMessage)
        {
            outputFilePath = null;
            errorMessage = string.Empty;

            try
            {
                var bodyBytes = Encoding.UTF8.GetBytes(requestJson);
                using (var content = new ByteArrayContent(bodyBytes))
                {
                    content.Headers.ContentType =
                        new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");

                    using (var request = new HttpRequestMessage(HttpMethod.Post, endpointUrl))
                    {
                        request.Headers.Authorization =
                            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiToken);
                        request.Content = content;

                        using (var response = SharedHttpClient.Value.SendAsync(request).Result)
                        {
                            var responseJson = response.Content.ReadAsStringAsync().Result;

                            if (!response.IsSuccessStatusCode)
                            {
                                errorMessage = "API request failed (" + (int)response.StatusCode + "): " + responseJson;
                                return false;
                            }

                            if (string.IsNullOrWhiteSpace(responseJson))
                            {
                                errorMessage = "API returned empty response.";
                                return false;
                            }

                            var base64Data = ExtractJsonStringValue(responseJson, "b64_json");
                            if (!string.IsNullOrWhiteSpace(base64Data))
                                return SaveBase64Image(base64Data, format, out outputFilePath, out errorMessage);

                            var imageUrl = ExtractJsonStringValue(responseJson, "url");
                            if (!string.IsNullOrWhiteSpace(imageUrl))
                                return LegacyDownloadFromUrl(imageUrl, format, apiToken, out outputFilePath, out errorMessage);

                            errorMessage = "No image data found in API response.";
                            return false;
                        }
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                errorMessage = "API request failed (network): " + ex.Message;
                return false;
            }
            catch (TaskCanceledException ex)
            {
                errorMessage = "API request timed out: " + ex.Message;
                return false;
            }
            catch (AggregateException ae)
            {
                var inner = ae.InnerException ?? ae;
                if (inner is TaskCanceledException)
                    errorMessage = "API request timed out: " + inner.Message;
                else
                    errorMessage = "API request failed: " + inner.Message;
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "Image generation failed: " + ex.Message;
                return false;
            }
        }

        private static bool LegacyDownloadFromUrl(string url, string format, string apiToken,
            out string filePath, out string error)
        {
            filePath = null;
            error = string.Empty;

            try
            {
                var ext = "." + (format ?? "png").TrimStart('.');
                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw", "AiImages");
                Directory.CreateDirectory(tempDir);
                filePath = Path.Combine(tempDir,
                    "ai_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" +
                    Guid.NewGuid().ToString("N").Substring(0, 6) + ext);

                using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                {
                    if (!string.IsNullOrWhiteSpace(apiToken))
                        request.Headers.Authorization =
                            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiToken);

                    using (var response = SharedHttpClient.Value.SendAsync(request).Result)
                    {
                        response.EnsureSuccessStatusCode();
                        var bytes = response.Content.ReadAsByteArrayAsync().Result;
                        File.WriteAllBytes(filePath, bytes);
                    }
                }

                if (File.Exists(filePath) && new FileInfo(filePath).Length > 0)
                    return true;

                error = "Image download failed.";
                return false;
            }
            catch (Exception ex)
            {
                error = "Image download failed: " + ex.Message;
                return false;
            }
        }


        /// <summary>
        /// Returns a file path to the ORIGINAL image behind the selected picture,
        /// at full resolution and without any recompression. The bytes are pulled
        /// straight from the embedded media inside the .pptx (via a snapshot), so
        /// resolution and clarity match the source — not PowerPoint's downscaled
        /// on-screen render. If the picture is cropped, only the visible region is
        /// returned (cropped losslessly from the original pixels). Falls back to a
        /// shape export for linked / non-embedded images.
        /// </summary>
        public static string GetSelectedOriginalImageFile(out string errorMessage)
        {
            errorMessage = string.Empty;
            string snapshotPath = null;
            try
            {
                dynamic app = Globals.ThisAddIn?.Application;
                if (app == null) { errorMessage = "未能获取 PowerPoint 应用实例。"; return null; }

                dynamic selection = null;
                try { selection = app.ActiveWindow?.Selection; }
                catch { errorMessage = "未能获取选区。"; return null; }

                if (selection == null || selection.Type != 2 || selection.ShapeRange.Count < 1)
                {
                    errorMessage = "请先选择一张图片。";
                    return null;
                }

                dynamic shape = selection.ShapeRange[1];

                if (shape.Type == 6) // msoGroup -> first picture child
                {
                    dynamic found = null;
                    try
                    {
                        foreach (var child in shape.GroupItems)
                        {
                            if (child.Type == 13 || child.Type == 11) { found = child; break; }
                        }
                    }
                    catch { }
                    if (found == null) { errorMessage = "选中的组合中没有图片。"; return null; }
                    shape = found;
                }
                else if (shape.Type != 13 && shape.Type != 11)
                {
                    errorMessage = "请先选择一张图片。";
                    return null;
                }

                // Snapshot the presentation, then extract the original embedded media bytes.
                string snapErr;
                if (!ImageReplacePipeline.TryCreatePresentationSnapshot(app, out snapshotPath, out snapErr))
                {
                    errorMessage = snapErr;
                    return null;
                }

                string originalPath;
                string extractErr;
                bool extracted = ImageReplacePipeline.TryExtractOriginalImageFromPptx(
                    shape, snapshotPath, out originalPath, out extractErr);

                if (!extracted || string.IsNullOrWhiteSpace(originalPath) || !File.Exists(originalPath))
                {
                    // Linked or non-embedded image: fall back to a high-quality shape export.
                    return ExportSelectedShapeFallback(shape, out errorMessage);
                }

                // If the picture is cropped, return only the visible region — cut from the
                // original pixels (no resampling), so we still upload the un-compressed source.
                float cl, ct, cr, cb;
                if (ImageReplacePipeline.TryGetPictureCropValues(shape, out cl, out ct, out cr, out cb)
                    && (Math.Abs(cl) > 0.01f || Math.Abs(ct) > 0.01f || Math.Abs(cr) > 0.01f || Math.Abs(cb) > 0.01f))
                {
                    float w = (float)shape.Width;
                    float h = (float)shape.Height;
                    string cropped = TryCropOriginalToVisibleRegion(originalPath, w, h, cl, ct, cr, cb);
                    if (!string.IsNullOrWhiteSpace(cropped))
                        return cropped;
                }

                return originalPath;
            }
            catch (Exception ex)
            {
                errorMessage = "获取原图失败：" + ex.Message;
                return null;
            }
            finally
            {
                try { if (!string.IsNullOrWhiteSpace(snapshotPath) && File.Exists(snapshotPath)) File.Delete(snapshotPath); }
                catch { }
            }
        }

        /// <summary>
        /// Crops the visible region out of the original image pixels, 1:1 (no resampling),
        /// and saves it losslessly as PNG. Crop edges are in points and map to fractions of
        /// the full (uncropped) image, matching the geometry model used by the color pipeline.
        /// Returns null on failure (caller then uses the full original).
        /// </summary>
        private static string TryCropOriginalToVisibleRegion(
            string originalPath, float shapeW, float shapeH, float cl, float ct, float cr, float cb)
        {
            try
            {
                float fullW = shapeW + cl + cr;
                float fullH = shapeH + ct + cb;
                if (fullW <= 0f || fullH <= 0f)
                    return null;

                double fracL = Math.Max(0.0, cl / fullW);
                double fracT = Math.Max(0.0, ct / fullH);
                double fracR = Math.Max(0.0, cr / fullW);
                double fracB = Math.Max(0.0, cb / fullH);
                double keepW = 1.0 - fracL - fracR;
                double keepH = 1.0 - fracT - fracB;
                if (keepW <= 0.0 || keepH <= 0.0)
                    return null;

                using (var src = Image.FromFile(originalPath))
                {
                    int ow = src.Width;
                    int oh = src.Height;
                    int x = (int)Math.Round(fracL * ow);
                    int y = (int)Math.Round(fracT * oh);
                    int cw = (int)Math.Round(keepW * ow);
                    int ch = (int)Math.Round(keepH * oh);

                    x = Math.Max(0, Math.Min(x, ow - 1));
                    y = Math.Max(0, Math.Min(y, oh - 1));
                    cw = Math.Max(1, Math.Min(cw, ow - x));
                    ch = Math.Max(1, Math.Min(ch, oh - y));

                    // Effectively uncropped — keep the original file as-is.
                    if (x == 0 && y == 0 && cw == ow && ch == oh)
                        return originalPath;

                    using (var bmp = new Bitmap(cw, ch, PixelFormat.Format32bppArgb))
                    {
                        using (var g = Graphics.FromImage(bmp))
                        {
                            // Source and destination rects are identical in size, so this is a
                            // straight pixel copy — no scaling, no quality loss.
                            g.InterpolationMode = InterpolationMode.NearestNeighbor;
                            g.PixelOffsetMode = PixelOffsetMode.Half;
                            g.DrawImage(src,
                                new Rectangle(0, 0, cw, ch),
                                new Rectangle(x, y, cw, ch),
                                GraphicsUnit.Pixel);
                        }

                        var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw", "ImgToImg");
                        Directory.CreateDirectory(tempDir);
                        var outPath = Path.Combine(tempDir,
                            "src_crop_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png");
                        bmp.Save(outPath, ImageFormat.Png);
                        return outPath;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Fallback for linked / non-embedded pictures: export the shape to a PNG at its
        /// native pixel size. Used only when the original cannot be pulled from the .pptx.
        /// </summary>
        private static string ExportSelectedShapeFallback(dynamic shape, out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw", "ImgToImg");
                Directory.CreateDirectory(tempDir);
                var tempPath = Path.Combine(tempDir,
                    "src_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png");

                float w = (float)shape.Width;
                float h = (float)shape.Height;
                try { shape.Export(tempPath, 2, (int)Math.Round(w), (int)Math.Round(h)); }
                catch { shape.Export(tempPath, 2); }

                if (File.Exists(tempPath) && new FileInfo(tempPath).Length > 0)
                    return tempPath;

                errorMessage = "未能导出选中的图片。";
                return null;
            }
            catch (Exception ex)
            {
                errorMessage = "导出选中图片失败：" + ex.Message;
                return null;
            }
        }

        /// <summary>
        /// Gets the dimensions of the selected picture shape. Returns (0,0) on failure.
        /// </summary>

        public static void GetSelectedImageDimensions(out int width, out int height)
        {
            width = 0;
            height = 0;
            try
            {
                dynamic app = Globals.ThisAddIn?.Application;
                if (app == null) return;

                dynamic selection = null;
                try { selection = app.ActiveWindow?.Selection; }
                catch { return; }

                if (selection == null || selection.Type != 2) return;
                if (selection.ShapeRange.Count < 1) return;

                dynamic shape = selection.ShapeRange[1];
                if (shape.Type == 6) // group
                {
                    try
                    {
                        foreach (var child in shape.GroupItems)
                        {
                            if (child.Type == 13 || child.Type == 11)
                            {
                                shape = child;
                                break;
                            }
                        }
                    }
                    catch { return; }
                }

                if (shape.Type != 13 && shape.Type != 11) return;

                width = (int)Math.Round((float)shape.Width);
                height = (int)Math.Round((float)shape.Height);
            }
            catch { }
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
