using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class GlobalSettingsWebDialog : WebViewDialogBase
    {
        private readonly AiImageGlobalSettings _settings;

        public GlobalSettingsWebDialog(AiImageGlobalSettings settings)
        {
            _settings = settings ?? AiImageService.CreateDefaultGlobalSettings();
            Text = "文生图 全局设置";
            InitWebView(500, 420, "global-settings.html");
        }

        protected override void OnWebViewReady()
        {
            PostMessageToWeb("init", new Dictionary<string, object>
            {
                { "override", _settings.OverridePerModel },
                { "previewCount", AiImageService.ClampModelPreviewCount(_settings.ModelPreviewCount) },
                { "defaultWidth", ClampInt(_settings.DefaultWidth, 1, 4096) },
                { "defaultHeight", ClampInt(_settings.DefaultHeight, 1, 4096) },
                { "quality", NormalizeChoice(_settings.DefaultQuality, "auto", "auto", "low", "medium", "high") },
                { "format", NormalizeChoice(_settings.DefaultFormat, "png", "png", "jpeg", "webp") }
            });
        }

        protected override void HandleWebMessage(
            string action, Dictionary<string, object> payload)
        {
            if (string.Equals(action, "save", StringComparison.Ordinal))
            {
                _settings.OverridePerModel = GetBool(payload, "override");
                _settings.ModelPreviewCount = AiImageService.ClampModelPreviewCount(
                    GetInt(payload, "previewCount", 5));
                _settings.DefaultWidth = ClampInt(GetInt(payload, "defaultWidth", 1024), 1, 4096);
                _settings.DefaultHeight = ClampInt(GetInt(payload, "defaultHeight", 1024), 1, 4096);
                _settings.DefaultQuality = NormalizeChoice(
                    GetString(payload, "quality"), "auto", "auto", "low", "medium", "high");
                _settings.DefaultFormat = NormalizeChoice(
                    GetString(payload, "format"), "png", "png", "jpeg", "webp");
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            if (string.Equals(action, "cancel", StringComparison.Ordinal))
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }
    }
}
