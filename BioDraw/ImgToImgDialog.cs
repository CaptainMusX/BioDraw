using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class ImgToImgDialog : Form
    {
        private readonly AiImageApiSettings _settings;
        private readonly List<AiImageApiSettings> _allSettings;
        private TextBox _promptBox;
        private ComboBox _widthCombo;
        private ComboBox _heightCombo;
        private ComboBox _qualityCombo;
        private ComboBox _formatCombo;
        private Button _generateButton;
        private Button _cancelButton;
        private Label _statusLabel;
        private bool _isGenerating;
        private int _initialWidth;
        private int _initialHeight;

        private static readonly string[] AiSizeOptions =
        {
            "512", "768", "1024", "1536", "1792", "1920",
            "2048", "2160", "2560", "3072", "3840", "4096"
        };

        private static readonly string[] QualityOptions = { "auto (自动)", "low (低)", "medium (中)", "high (高)" };
        private static readonly string[] QualityValues = { "auto", "low", "medium", "high" };
        private static readonly string[] FormatOptions = { "PNG", "JPEG", "WebP" };
        private static readonly string[] FormatValues = { "png", "jpeg", "webp" };

        public ImgToImgDialog(AiImageApiSettings settings, List<AiImageApiSettings> allSettings,
            int initialWidth, int initialHeight, string initialQuality, string initialFormat)
        {
            _settings = settings ?? AiImageService.CreateDefaultSettings();
            _allSettings = allSettings ?? new List<AiImageApiSettings> { _settings };
            _initialWidth = initialWidth > 0 ? initialWidth : (_settings.DefaultWidth > 0 ? _settings.DefaultWidth : 1024);
            _initialHeight = initialHeight > 0 ? initialHeight : (_settings.DefaultHeight > 0 ? _settings.DefaultHeight : 1024);
            InitializeForm(initialQuality, initialFormat);
        }

        private void InitializeForm(string initialQuality, string initialFormat)
        {
            Text = "图生图 - " + (_settings.DisplayName ?? "GPT-Image-2");
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Regular, GraphicsUnit.Point);
            BackColor = Color.FromArgb(244, 247, 252);
            ForeColor = Color.FromArgb(32, 41, 57);
            AutoScaleMode = AutoScaleMode.Dpi;
            MinimizeBox = false;
            MaximizeBox = true;
            MinimumSize = new Size(640, 480);
            ClientSize = new Size(720, 560);
            Icon = SystemIcons.Application;

            var promptLabel = new Label
            {
                Text = "提示词",
                TextAlign = ContentAlignment.MiddleLeft
            };

            _promptBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            var widthLabel = new Label
            {
                Text = "宽度",
                TextAlign = ContentAlignment.MiddleRight
            };
            _widthCombo = CreateEditableCombo(AiSizeOptions, GetClosestIndex(_initialWidth));

            var heightLabel = new Label
            {
                Text = "高度",
                TextAlign = ContentAlignment.MiddleRight
            };
            _heightCombo = CreateEditableCombo(AiSizeOptions, GetClosestIndex(_initialHeight));

            var qualityLabel = new Label
            {
                Text = "质量",
                TextAlign = ContentAlignment.MiddleRight
            };
            _qualityCombo = CreateCombo(QualityOptions, GetQualityIndex(initialQuality ?? _settings.DefaultQuality));

            var formatLabel = new Label
            {
                Text = "格式",
                TextAlign = ContentAlignment.MiddleRight
            };
            _formatCombo = CreateCombo(FormatOptions, GetFormatIndex(initialFormat ?? _settings.DefaultFormat));

            _statusLabel = new Label
            {
                Text = "就绪",
                ForeColor = Color.FromArgb(140, 149, 166),
                TextAlign = ContentAlignment.MiddleLeft
            };

            _generateButton = new Button
            {
                Text = "生图",
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand
            };
            _generateButton.FlatAppearance.BorderSize = 1;
            _generateButton.FlatAppearance.BorderColor = Color.FromArgb(24, 118, 242);
            _generateButton.BackColor = Color.FromArgb(24, 118, 242);
            _generateButton.ForeColor = Color.White;
            _generateButton.Click += OnGenerateClick;

            _cancelButton = new Button
            {
                Text = "关闭",
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand
            };
            _cancelButton.FlatAppearance.BorderSize = 1;
            _cancelButton.FlatAppearance.BorderColor = Color.FromArgb(189, 198, 213);
            _cancelButton.BackColor = Color.White;
            _cancelButton.ForeColor = Color.FromArgb(43, 52, 69);
            _cancelButton.Click += (s, e) => Close();

            Controls.Add(promptLabel);
            Controls.Add(_promptBox);
            Controls.Add(widthLabel);
            Controls.Add(_widthCombo);
            Controls.Add(heightLabel);
            Controls.Add(_heightCombo);
            Controls.Add(qualityLabel);
            Controls.Add(_qualityCombo);
            Controls.Add(formatLabel);
            Controls.Add(_formatCombo);
            Controls.Add(_statusLabel);
            Controls.Add(_generateButton);
            Controls.Add(_cancelButton);

            form_Resize();
            Resize += (_, __) => form_Resize();
            AcceptButton = _generateButton;
        }

        private void form_Resize()
        {
            var margin = 24;
            var labelHeight = 28;
            var inputHeight = 32;
            var rowGap = 16;
            var buttonWidth = 128;
            var buttonHeight = 40;
            var fieldGap = 8;
            var labelWidth = TextRenderer.MeasureText("宽度", Font).Width + 8;

            var top = margin;
            var rightEdge = ClientSize.Width - margin;

            var promptLabel = (Label)Controls[0];
            promptLabel.Location = new Point(margin, top);
            promptLabel.Size = new Size(80, labelHeight);

            top += labelHeight + 4;

            var promptHeight = Math.Max(140, ClientSize.Height - margin - buttonHeight - margin - (inputHeight + rowGap) * 2 - labelHeight - 4 - rowGap);

            _promptBox.Location = new Point(margin, top);
            _promptBox.Size = new Size(rightEdge - margin, promptHeight);

            top += promptHeight + rowGap;

            var fieldX = margin + labelWidth + fieldGap;
            var totalFieldWidth = rightEdge - fieldX;
            var halfWidth = totalFieldWidth / 2;
            var comboWidth = Math.Max(140, halfWidth - labelWidth - fieldGap - 16);

            var col2LabelX = fieldX + comboWidth + 16;
            var col2FieldX = col2LabelX + labelWidth + fieldGap;

            var widthLabel = (Label)Controls[2];
            widthLabel.Location = new Point(margin, top);
            widthLabel.Size = new Size(labelWidth, inputHeight);
            _widthCombo.Location = new Point(fieldX, top);
            _widthCombo.Size = new Size(comboWidth, inputHeight);

            var heightLabel = (Label)Controls[4];
            heightLabel.Location = new Point(col2LabelX, top);
            heightLabel.Size = new Size(labelWidth, inputHeight);
            _heightCombo.Location = new Point(col2FieldX, top);
            _heightCombo.Size = new Size(comboWidth, inputHeight);

            top += inputHeight + rowGap;

            var qualityLabel = (Label)Controls[6];
            qualityLabel.Location = new Point(margin, top);
            qualityLabel.Size = new Size(labelWidth, inputHeight);
            _qualityCombo.Location = new Point(fieldX, top);
            _qualityCombo.Size = new Size(comboWidth, inputHeight);

            var formatLabel = (Label)Controls[8];
            formatLabel.Location = new Point(col2LabelX, top);
            formatLabel.Size = new Size(labelWidth, inputHeight);
            _formatCombo.Location = new Point(col2FieldX, top);
            _formatCombo.Size = new Size(comboWidth, inputHeight);

            var bottomY = ClientSize.Height - margin - buttonHeight;
            _statusLabel.Location = new Point(margin, bottomY + (buttonHeight - labelHeight) / 2);
            _statusLabel.Size = new Size(rightEdge - margin - (buttonWidth * 2) - 32, labelHeight);

            _cancelButton.Location = new Point(rightEdge - buttonWidth * 2 - 16, bottomY);
            _cancelButton.Size = new Size(buttonWidth, buttonHeight);
            _generateButton.Location = new Point(rightEdge - buttonWidth, bottomY);
            _generateButton.Size = new Size(buttonWidth, buttonHeight);
        }

        private static ComboBox CreateEditableCombo(string[] items, int selectedIndex)
        {
            var combo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDown,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(32, 41, 57),
                FlatStyle = FlatStyle.Flat
            };
            combo.Items.AddRange(items);
            if (selectedIndex >= 0 && selectedIndex < items.Length)
                combo.SelectedIndex = selectedIndex;
            return combo;
        }

        private static ComboBox CreateCombo(string[] items, int selectedIndex)
        {
            var combo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(32, 41, 57),
                FlatStyle = FlatStyle.Flat
            };
            combo.Items.AddRange(items);
            combo.SelectedIndex = selectedIndex >= 0 && selectedIndex < items.Length ? selectedIndex : 0;
            return combo;
        }

        private int GetClosestIndex(int value)
        {
            for (int i = 0; i < AiSizeOptions.Length; i++)
            {
                if (int.TryParse(AiSizeOptions[i], NumberStyles.Integer, CultureInfo.InvariantCulture, out int v)
                    && v >= value)
                    return i;
            }
            return AiSizeOptions.Length - 1;
        }

        private int GetQualityIndex(string value)
        {
            for (int i = 0; i < QualityValues.Length; i++)
                if (string.Equals(QualityValues[i], value, StringComparison.OrdinalIgnoreCase))
                    return i;
            return 0;
        }

        private int GetFormatIndex(string value)
        {
            for (int i = 0; i < FormatValues.Length; i++)
                if (string.Equals(FormatValues[i], value, StringComparison.OrdinalIgnoreCase))
                    return i;
            return 0;
        }

        private void OnGenerateClick(object sender, EventArgs e)
        {
            if (_isGenerating)
                return;

            var prompt = (_promptBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(prompt))
            {
                MessageBox.Show("请输入提示词。", "BioDraw", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _promptBox.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(_settings.ApiToken))
            {
                var result = MessageBox.Show(
                    "尚未设置 API 令牌。\n\n是否现在打开设置？",
                    "BioDraw",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                    ShowSettingsDialog();
                return;
            }

            // Capture source image base64 before spawning thread
            var sourceBase64 = AiImageService.GetSelectedImageBase64();
            if (string.IsNullOrWhiteSpace(sourceBase64))
            {
                _statusLabel.Text = "未能获取选中的图片，请重新选择。";
                _statusLabel.ForeColor = Color.FromArgb(220, 53, 69);
                return;
            }

            SetGeneratingState(true);

            var widthText = (_widthCombo.Text ?? string.Empty).Trim();
            var heightText = (_heightCombo.Text ?? string.Empty).Trim();
            int.TryParse(widthText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int width);
            int.TryParse(heightText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int height);
            if (width < 1) width = _initialWidth;
            if (height < 1) height = _initialHeight;

            var quality = QualityValues[_qualityCombo.SelectedIndex];
            var format = FormatValues[_formatCombo.SelectedIndex];

            var thread = new Thread(() =>
            {
                string outputPath;
                string error;
                var success = AiImageService.TryGenerateImageFromImage(_settings, prompt, width, height, quality, format,
                    sourceBase64, out outputPath, out error);

                BeginInvoke(new Action(() =>
                {
                    SetGeneratingState(false);

                    if (success && File.Exists(outputPath))
                    {
                        var insertError = TryInsertToSlide(outputPath);
                        if (string.IsNullOrEmpty(insertError))
                        {
                            _statusLabel.Text = "已插入幻灯片";
                            _statusLabel.ForeColor = Color.FromArgb(40, 167, 69);
                            Close();
                        }
                        else
                        {
                            _statusLabel.Text = insertError;
                            _statusLabel.ForeColor = Color.FromArgb(220, 53, 69);
                        }
                    }
                    else
                    {
                        _statusLabel.Text = error ?? "生成失败。";
                        _statusLabel.ForeColor = Color.FromArgb(220, 53, 69);
                    }
                }));
            });

            thread.IsBackground = true;
            thread.Start();
        }

        private void SetGeneratingState(bool generating)
        {
            _isGenerating = generating;
            _generateButton.Enabled = !generating;
            _promptBox.Enabled = !generating;
            _widthCombo.Enabled = !generating;
            _heightCombo.Enabled = !generating;
            _qualityCombo.Enabled = !generating;
            _formatCombo.Enabled = !generating;

            if (generating)
            {
                _statusLabel.Text = "正在生成，请稍候...";
                _statusLabel.ForeColor = Color.FromArgb(24, 118, 242);
            }
        }

        private static string TryInsertToSlide(string filePath)
        {
            try
            {
                dynamic app = Globals.ThisAddIn?.Application;
                if (app == null)
                    return "未能获取 PowerPoint 应用实例。";

                dynamic slide = null;
                try { slide = app.ActiveWindow?.View?.Slide; }
                catch { }

                if (slide == null)
                    return "请先切换到普通编辑视图。";

                dynamic newShape = slide.Shapes.AddPicture(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    0f, 0f, -1f, -1f);

                var pageSetup = app.ActivePresentation?.PageSetup;
                if (pageSetup != null)
                {
                    float slideWidth = (float)pageSetup.SlideWidth;
                    float slideHeight = (float)pageSetup.SlideHeight;
                    float maxWidth = slideWidth * 0.8f;
                    float maxHeight = slideHeight * 0.8f;

                    float picWidth = (float)newShape.Width;
                    float picHeight = (float)newShape.Height;

                    if (picWidth > maxWidth || picHeight > maxHeight)
                    {
                        float scale = Math.Min(maxWidth / picWidth, maxHeight / picHeight);
                        newShape.LockAspectRatio = -1;
                        newShape.Width = picWidth * scale;
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

        private void ShowSettingsDialog()
        {
            using (var dialog = new AiImageSettingsDialog(_settings, _allSettings))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    AiImageService.SaveSettings(_allSettings, AiImageService.LoadGlobalSettings());
            }
        }
    }
}
