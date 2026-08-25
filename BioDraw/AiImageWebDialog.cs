using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class AiImageWebDialog : WebViewDialogBase
    {
        private readonly AiImageApiSettings _settings;
        private readonly List<AiImageApiSettings> _allSettings;
        private readonly int _initialWidth;
        private readonly int _initialHeight;
        private readonly string _initialQuality;
        private readonly string _initialFormat;
        private bool _isGenerating;

        public AiImageWebDialog(
            AiImageApiSettings settings,
            List<AiImageApiSettings> allSettings,
            int initialWidth,
            int initialHeight,
            string initialQuality,
            string initialFormat)
        {
            _settings = settings ?? AiImageService.CreateDefaultSettings();
            _allSettings = allSettings ?? new List<AiImageApiSettings> { _settings };
            _initialWidth = ClampInt(
                initialWidth > 0 ? initialWidth : _settings.DefaultWidth, 1, 4096);
            _initialHeight = ClampInt(
                initialHeight > 0 ? initialHeight : _settings.DefaultHeight, 1, 4096);
            _initialQuality = NormalizeChoice(
                initialQuality ?? _settings.DefaultQuality, "auto", "auto", "low", "medium", "high");
            _initialFormat = NormalizeChoice(
                initialFormat ?? _settings.DefaultFormat, "png", "png", "jpeg", "webp");
            Text = "素材自绘 - " + (_settings.DisplayName ?? "GPT-Image-2");
            InitWebView(620, 560, "image-gen.html");
        }

        protected override void OnWebViewReady()
        {
            PostMessageToWeb("init", new Dictionary<string, object>
            {
                { "title", "素材自绘 - " + (_settings.DisplayName ?? "GPT-Image-2") },
                { "width", _initialWidth },
                { "height", _initialHeight },
                { "quality", _initialQuality },
                { "format", _initialFormat },
                { "showLock", false }
            });
        }

        protected override void HandleWebMessage(
            string action, Dictionary<string, object> payload)
        {
            if (string.Equals(action, "generate", StringComparison.Ordinal))
            {
                DoGenerate(payload);
                return;
            }

            if (string.Equals(action, "cancel", StringComparison.Ordinal)) Close();
        }

        private void DoGenerate(Dictionary<string, object> payload)
        {
            if (_isGenerating) return;

            var prompt = GetString(payload, "prompt").Trim();
            if (string.IsNullOrWhiteSpace(prompt))
            {
                SendStatus("请输入提示词。", "error");
                return;
            }
            if (prompt.Length > 20000)
            {
                SendStatus("提示词过长，请缩短后重试。", "error");
                return;
            }

            if (string.IsNullOrWhiteSpace(_settings.ApiToken))
            {
                var settingsSaved = false;
                var choice = MessageBox.Show(
                    "尚未设置 API 令牌。\n\n是否现在打开设置？",
                    "BioDraw",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (choice == DialogResult.Yes)
                {
                    using (var dialog = new ModelSettingsWebDialog(_settings, _allSettings))
                    {
                        if (dialog.ShowDialog(this) == DialogResult.OK)
                        {
                            AiImageService.SaveSettings(
                                _allSettings, AiImageService.LoadGlobalSettings());
                            settingsSaved = true;
                        }
                    }
                }
                var tokenReady = settingsSaved && !string.IsNullOrWhiteSpace(_settings.ApiToken);
                SendStatus(tokenReady
                    ? "API 设置已保存，请再次点击生图。"
                    : "请先配置 API 令牌。",
                    tokenReady ? "ready" : "error");
                return;
            }

            var width = ClampInt(GetInt(payload, "width", _initialWidth), 1, 4096);
            var height = ClampInt(GetInt(payload, "height", _initialHeight), 1, 4096);
            var quality = NormalizeChoice(
                GetString(payload, "quality"), _initialQuality, "auto", "low", "medium", "high");
            var format = NormalizeChoice(
                GetString(payload, "format"), _initialFormat, "png", "jpeg", "webp");

            _isGenerating = true;
            SendStatus("正在生成，请稍候...", "generating");

            Task.Run(() =>
            {
                string outputPath;
                string error;
                var success = AiImageService.TryGenerateImage(
                    _settings, prompt, width, height, quality, format,
                    out outputPath, out error);

                TryBeginInvoke(() =>
                {
                    _isGenerating = false;
                    if (success && File.Exists(outputPath))
                    {
                        var insertError = TryInsertToSlide(outputPath);
                        if (string.IsNullOrEmpty(insertError))
                        {
                            PostMessageToWeb("done");
                            CloseAfter(350);
                        }
                        else
                        {
                            SendStatus(insertError, "error");
                        }
                    }
                    else
                    {
                        SendStatus(error ?? "生成失败。", "error");
                    }
                });
            });
        }

        private void SendStatus(string text, string type)
        {
            PostMessageToWeb("status", new Dictionary<string, object>
            {
                { "text", text ?? string.Empty },
                { "type", type ?? "ready" }
            });
        }

        private static string TryInsertToSlide(string filePath)
        {
            try
            {
                dynamic app = Globals.ThisAddIn == null
                    ? null
                    : Globals.ThisAddIn.Application;
                if (app == null) return "未能获取 PowerPoint 应用实例。";

                dynamic slide = null;
                try { slide = app.ActiveWindow == null ? null : app.ActiveWindow.View.Slide; }
                catch { }
                if (slide == null) return "请先切换到普通编辑视图。";

                dynamic newShape = slide.Shapes.AddPicture(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    0f, 0f, -1f, -1f);

                var pageSetup = app.ActivePresentation == null
                    ? null
                    : app.ActivePresentation.PageSetup;
                if (pageSetup != null)
                {
                    float slideWidth = (float)pageSetup.SlideWidth;
                    float slideHeight = (float)pageSetup.SlideHeight;
                    float maxWidth = slideWidth * 0.8f;
                    float maxHeight = slideHeight * 0.8f;
                    float pictureWidth = (float)newShape.Width;
                    float pictureHeight = (float)newShape.Height;

                    if (pictureWidth > maxWidth || pictureHeight > maxHeight)
                    {
                        float scale = Math.Min(
                            maxWidth / pictureWidth, maxHeight / pictureHeight);
                        newShape.LockAspectRatio = -1;
                        newShape.Width = pictureWidth * scale;
                    }

                    newShape.Left = (slideWidth - (float)newShape.Width) / 2f;
                    newShape.Top = (slideHeight - (float)newShape.Height) / 2f;
                }

                newShape.Select();
                return null;
            }
            catch (Exception ex)
            {
                return "插入幻灯片失败：" + ex.Message;
            }
        }
    }
}
