using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class ModelSettingsWebDialog : WebViewDialogBase
    {
        private const long MaxIconFileBytes = 10L * 1024L * 1024L;

        private readonly AiImageApiSettings _settings;
        private readonly List<AiImageApiSettings> _allSettings;
        private string _pendingIconPath;

        public ModelSettingsWebDialog(
            AiImageApiSettings settings, List<AiImageApiSettings> allSettings)
        {
            _settings = settings ?? AiImageService.CreateDefaultSettings();
            _allSettings = allSettings ?? new List<AiImageApiSettings>();
            _pendingIconPath = _settings.IconPath ?? string.Empty;
            Text = "模型设置 - " + (_settings.DisplayName ?? "GPT-Image-2");
            InitWebView(560, 580, "model-settings.html");
        }

        protected override void OnWebViewReady()
        {
            var payload = new Dictionary<string, object>
            {
                { "displayName", _settings.DisplayName ?? string.Empty },
                { "apiToken", _settings.ApiToken ?? string.Empty },
                { "endpointUrl", _settings.EndpointUrl ?? string.Empty },
                { "model", _settings.Model ?? string.Empty },
                { "defaultWidth", ClampInt(_settings.DefaultWidth, 1, 4096) },
                { "defaultHeight", ClampInt(_settings.DefaultHeight, 1, 4096) },
                { "quality", NormalizeChoice(_settings.DefaultQuality, "auto", "auto", "low", "medium", "high") },
                { "format", NormalizeChoice(_settings.DefaultFormat, "png", "png", "jpeg", "webp") },
                { "resolution", NormalizeChoice(_settings.Resolution, "2k", "1k", "2k", "4k") },
                { "lockAspectRatio", _settings.LockAspectRatio },
                { "canDelete", _allSettings.Count > 1 }
            };

            var iconBase64 = TryCreateIconPreview(_pendingIconPath);
            if (!string.IsNullOrWhiteSpace(iconBase64))
                payload["iconBase64"] = iconBase64;

            PostMessageToWeb("init", payload);
        }

        protected override void HandleWebMessage(
            string action, Dictionary<string, object> payload)
        {
            switch (action)
            {
                case "save":
                    ApplySave(payload);
                    break;
                case "cancel":
                    DialogResult = DialogResult.Cancel;
                    Close();
                    break;
                case "delete":
                    DeleteSettings();
                    break;
                case "pickIcon":
                    PickIcon();
                    break;
            }
        }

        private void ApplySave(Dictionary<string, object> payload)
        {
            var token = GetString(payload, "apiToken").Trim();
            if (token.Length > 8192)
            {
                MessageBox.Show("API 令牌过长，请检查输入。", "BioDraw",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(token))
            {
                var choice = MessageBox.Show(
                    "API 令牌为空，将无法正常生图。\n\n是否继续保存？",
                    "BioDraw",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (choice != DialogResult.Yes) return;
            }

            _settings.DisplayName = Limit(GetString(payload, "displayName").Trim(), 100);
            _settings.ApiToken = token;
            _settings.EndpointUrl = Limit(GetString(payload, "endpointUrl").Trim(), 2048);
            _settings.Model = Limit(GetString(payload, "model").Trim(), 200);
            _settings.DefaultWidth = ClampInt(GetInt(payload, "defaultWidth", 1024), 1, 4096);
            _settings.DefaultHeight = ClampInt(GetInt(payload, "defaultHeight", 1024), 1, 4096);
            _settings.DefaultQuality = NormalizeChoice(
                GetString(payload, "quality"), "auto", "auto", "low", "medium", "high");
            _settings.DefaultFormat = NormalizeChoice(
                GetString(payload, "format"), "png", "png", "jpeg", "webp");
            _settings.Resolution = NormalizeChoice(
                GetString(payload, "resolution"), "2k", "1k", "2k", "4k");
            _settings.LockAspectRatio = GetBool(payload, "lockAspectRatio");
            _settings.IconPath = _pendingIconPath;

            if (!_allSettings.Contains(_settings)) _allSettings.Add(_settings);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void DeleteSettings()
        {
            if (_allSettings.Count <= 1)
            {
                MessageBox.Show("至少需要保留 1 个模型配置。", "BioDraw",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var name = _settings.DisplayName ?? _settings.Model ?? "当前模型";
            var choice = MessageBox.Show(
                "确认删除模型 \"" + name + "\" 吗？\n\n此操作不可撤销。",
                "BioDraw",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (choice != DialogResult.Yes) return;

            _allSettings.Remove(_settings);
            DialogResult = DialogResult.OK;
            Close();
        }

        private void PickIcon()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.ico";
                dialog.Title = "选择图标";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;

                var preview = TryCreateIconPreview(dialog.FileName);
                if (string.IsNullOrWhiteSpace(preview))
                {
                    MessageBox.Show("无法读取该图片，或图片文件超过 10 MB。", "BioDraw",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _pendingIconPath = dialog.FileName;
                PostMessageToWeb("iconChanged", new Dictionary<string, object>
                {
                    { "base64", preview }
                });
            }
        }

        private static string TryCreateIconPreview(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
                var info = new FileInfo(path);
                if (info.Length <= 0 || info.Length > MaxIconFileBytes) return null;

                using (var source = Image.FromFile(path))
                using (var preview = new Bitmap(128, 128, PixelFormat.Format32bppArgb))
                using (var graphics = Graphics.FromImage(preview))
                using (var stream = new MemoryStream())
                {
                    graphics.Clear(Color.Transparent);
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;

                    var scale = Math.Min(128f / source.Width, 128f / source.Height);
                    int width = Math.Max(1, (int)Math.Round(source.Width * scale));
                    int height = Math.Max(1, (int)Math.Round(source.Height * scale));
                    int x = (128 - width) / 2;
                    int y = (128 - height) / 2;
                    graphics.DrawImage(source, new Rectangle(x, y, width, height));

                    preview.Save(stream, ImageFormat.Png);
                    return Convert.ToBase64String(stream.ToArray());
                }
            }
            catch
            {
                return null;
            }
        }

        private static string Limit(string value, int maximumLength)
        {
            value = value ?? string.Empty;
            return value.Length <= maximumLength
                ? value
                : value.Substring(0, maximumLength);
        }
    }
}
