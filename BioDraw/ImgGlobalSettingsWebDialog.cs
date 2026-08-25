using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class ImgGlobalSettingsWebDialog : WebViewDialogBase
    {
        private readonly AiImageGlobalSettings _settings;

        public ImgGlobalSettingsWebDialog(AiImageGlobalSettings settings)
        {
            _settings = settings ?? AiImageService.CreateDefaultGlobalSettings();
            Text = "图生图 全局设置";
            InitWebView(500, 420, "img-global-settings.html");
        }

        protected override void OnWebViewReady()
        {
            PostMessageToWeb("init", new Dictionary<string, object>
            {
                { "override", _settings.ImgOverridePerModel },
                { "previewCount", AiImageService.ClampModelPreviewCount(_settings.ImgModelPreviewCount) },
                { "defaultWidth", ClampInt(_settings.ImgDefaultWidth, 1, 4096) },
                { "defaultHeight", ClampInt(_settings.ImgDefaultHeight, 1, 4096) },
                { "quality", NormalizeChoice(_settings.ImgDefaultQuality, "auto", "auto", "low", "medium", "high") },
                { "format", NormalizeChoice(_settings.ImgDefaultFormat, "png", "png", "jpeg", "webp") }
            });
        }

        protected override void HandleWebMessage(
            string action, Dictionary<string, object> payload)
        {
            if (string.Equals(action, "save", StringComparison.Ordinal))
            {
                _settings.ImgOverridePerModel = GetBool(payload, "override");
                _settings.ImgModelPreviewCount = AiImageService.ClampModelPreviewCount(
                    GetInt(payload, "previewCount", 5));
                _settings.ImgDefaultWidth = ClampInt(GetInt(payload, "defaultWidth", 1024), 1, 4096);
                _settings.ImgDefaultHeight = ClampInt(GetInt(payload, "defaultHeight", 1024), 1, 4096);
                _settings.ImgDefaultQuality = NormalizeChoice(
                    GetString(payload, "quality"), "auto", "auto", "low", "medium", "high");
                _settings.ImgDefaultFormat = NormalizeChoice(
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
