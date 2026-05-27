using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class AiGlobalSettingsDialog : Form
    {
        private readonly AiImageGlobalSettings _settings;
        private CheckBox _overrideCheck;
        private NumericUpDown _previewCountNum;
        private TrackBar _previewCountSlider;
        private ComboBox _widthCombo;
        private ComboBox _heightCombo;
        private ComboBox _qualityCombo;
        private ComboBox _formatCombo;

        private static readonly string[] AiSizeOptions =
        {
            "512", "768", "1024", "1536", "1792", "1920",
            "2048", "2160", "2560", "3072", "3840", "4096"
        };

        private static readonly string[] QualityOptions = { "auto (自动)", "low (低)", "medium (中)", "high (高)" };
        private static readonly string[] QualityValues = { "auto", "low", "medium", "high" };
        private static readonly string[] FormatOptions = { "PNG", "JPEG", "WebP" };
        private static readonly string[] FormatValues = { "png", "jpeg", "webp" };

        public AiGlobalSettingsDialog(AiImageGlobalSettings settings)
        {
            _settings = settings ?? AiImageService.CreateDefaultGlobalSettings();
            InitializeForm();
        }

        private void InitializeForm()
        {
            Text = "全局设置";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Regular, GraphicsUnit.Point);
            BackColor = Color.FromArgb(244, 247, 252);
            ForeColor = Color.FromArgb(32, 41, 57);
            AutoScaleMode = AutoScaleMode.Dpi;
            MaximizeBox = true;
            MinimizeBox = false;
            MinimumSize = new Size(500, 420);
            ClientSize = new Size(560, 460);
            Icon = SystemIcons.Application;

            const int margin = 24;
            const int labelWidth = 72;
            const int fieldGap = 12;
            const int rowHeight = 32;
            const int rowGap = 16;
            // Field width must fit "auto (自动)" without clipping — at least 160px
            const int fieldWidth = 180;

            _overrideCheck = new CheckBox
            {
                Text = "通用配置优先于独立配置",
                AutoSize = true,
                Checked = _settings.OverridePerModel,
                ForeColor = Color.FromArgb(43, 52, 69)
            };

            var lblPreviewCount = new Label { Text = "预览数量", TextAlign = ContentAlignment.MiddleRight };

            _previewCountNum = new NumericUpDown
            {
                Minimum = 1,
                Maximum = AiImageService.AiModelButtonCount,
                DecimalPlaces = 0,
                Value = Math.Max(1, Math.Min(AiImageService.AiModelButtonCount, _settings.ModelPreviewCount)),
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Right,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(32, 41, 57),
                Width = 64
            };
            _previewCountNum.ValueChanged += OnPreviewNumChanged;

            _previewCountSlider = new TrackBar
            {
                Minimum = 1,
                Maximum = AiImageService.AiModelButtonCount,
                TickFrequency = 1,
                SmallChange = 1,
                LargeChange = 1,
                AutoSize = false,
                Value = Math.Max(1, Math.Min(AiImageService.AiModelButtonCount, _settings.ModelPreviewCount))
            };
            _previewCountSlider.ValueChanged += OnPreviewSliderChanged;

            var lblWidth = new Label { Text = "宽度", TextAlign = ContentAlignment.MiddleRight };
            _widthCombo = CreateEditableCombo(AiSizeOptions, GetClosestSizeIndex(_settings.DefaultWidth));
            _widthCombo.Text = _settings.DefaultWidth.ToString(CultureInfo.InvariantCulture);

            var lblHeight = new Label { Text = "高度", TextAlign = ContentAlignment.MiddleRight };
            _heightCombo = CreateEditableCombo(AiSizeOptions, GetClosestSizeIndex(_settings.DefaultHeight));
            _heightCombo.Text = _settings.DefaultHeight.ToString(CultureInfo.InvariantCulture);

            var lblQuality = new Label { Text = "质量", TextAlign = ContentAlignment.MiddleRight };
            _qualityCombo = CreateCombo(QualityOptions, GetQualityIndex(_settings.DefaultQuality));

            var lblFormat = new Label { Text = "格式", TextAlign = ContentAlignment.MiddleRight };
            _formatCombo = CreateCombo(FormatOptions, GetFormatIndex(_settings.DefaultFormat));

            var btnSave = new Button
            {
                Text = "保存",
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 1;
            btnSave.FlatAppearance.BorderColor = Color.FromArgb(24, 118, 242);
            btnSave.BackColor = Color.FromArgb(24, 118, 242);
            btnSave.ForeColor = Color.White;
            btnSave.Click += OnSaveClick;

            var btnCancel = new Button
            {
                Text = "取消",
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 1;
            btnCancel.FlatAppearance.BorderColor = Color.FromArgb(189, 198, 213);
            btnCancel.BackColor = Color.White;
            btnCancel.ForeColor = Color.FromArgb(43, 52, 69);
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            Controls.Add(_overrideCheck);
            Controls.Add(lblPreviewCount);
            Controls.Add(_previewCountNum);
            Controls.Add(_previewCountSlider);
            Controls.Add(lblWidth);
            Controls.Add(_widthCombo);
            Controls.Add(lblHeight);
            Controls.Add(_heightCombo);
            Controls.Add(lblQuality);
            Controls.Add(_qualityCombo);
            Controls.Add(lblFormat);
            Controls.Add(_formatCombo);
            Controls.Add(btnSave);
            Controls.Add(btnCancel);

            AcceptButton = btnSave;

            Resize += (_, __) => LayoutControls();
            LayoutControls();

            void LayoutControls()
            {
                var y = margin;
                var fieldX = margin + labelWidth + fieldGap;
                var dynFieldWidth = Math.Max(fieldWidth, ClientSize.Width - fieldX - margin);

                _overrideCheck.Location = new Point(margin, y);
                _overrideCheck.Size = new Size(ClientSize.Width - margin * 2, rowHeight);
                y += rowHeight + rowGap + 4;

                // Preview count row: label + num + slider
                lblPreviewCount.Location = new Point(margin, y);
                lblPreviewCount.Size = new Size(labelWidth, rowHeight);
                _previewCountNum.Location = new Point(fieldX, y + (rowHeight - _previewCountNum.Height) / 2);
                var sliderX = fieldX + 72;
                var sliderWidth = Math.Max(100, ClientSize.Width - sliderX - margin);
                _previewCountSlider.Location = new Point(sliderX, y);
                _previewCountSlider.Size = new Size(sliderWidth, rowHeight);
                y += rowHeight + rowGap;

                LayoutRow(lblWidth, _widthCombo, dynFieldWidth, ref y);
                LayoutRow(lblHeight, _heightCombo, dynFieldWidth, ref y);
                LayoutRow(lblQuality, _qualityCombo, dynFieldWidth, ref y);
                LayoutRow(lblFormat, _formatCombo, dynFieldWidth, ref y);

                var bottomY = ClientSize.Height - margin - 40;
                const int buttonWidth = 100;
                const int buttonHeight = 40;
                var rightEdge = ClientSize.Width - margin;
                btnSave.Location = new Point(rightEdge - buttonWidth, bottomY);
                btnSave.Size = new Size(buttonWidth, buttonHeight);
                btnCancel.Location = new Point(rightEdge - buttonWidth * 2 - 12, bottomY);
                btnCancel.Size = new Size(buttonWidth, buttonHeight);
            }

            void LayoutRow(Control label, Control field, int fw, ref int y2)
            {
                var fieldX = margin + labelWidth + fieldGap;
                label.Location = new Point(margin, y2);
                label.Size = new Size(labelWidth, rowHeight);
                field.Location = new Point(fieldX, y2);
                field.Size = new Size(fw, rowHeight);
                y2 += rowHeight + rowGap;
            }
        }

        private void OnPreviewNumChanged(object sender, EventArgs e)
        {
            var v = (int)_previewCountNum.Value;
            if (_previewCountSlider.Value != v)
                _previewCountSlider.Value = Math.Max(1, Math.Min(AiImageService.AiModelButtonCount, v));
        }

        private void OnPreviewSliderChanged(object sender, EventArgs e)
        {
            if (_previewCountNum.Value != _previewCountSlider.Value)
                _previewCountNum.Value = _previewCountSlider.Value;
        }

        private static int GetClosestSizeIndex(int value)
        {
            for (int i = 0; i < AiSizeOptions.Length; i++)
            {
                if (int.TryParse(AiSizeOptions[i], NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out int v) && v >= value)
                    return i;
            }
            return AiSizeOptions.Length - 1;
        }

        private static int GetQualityIndex(string value)
        {
            for (int i = 0; i < QualityValues.Length; i++)
                if (string.Equals(QualityValues[i], value, StringComparison.OrdinalIgnoreCase))
                    return i;
            return 0;
        }

        private static int GetFormatIndex(string value)
        {
            for (int i = 0; i < FormatValues.Length; i++)
                if (string.Equals(FormatValues[i], value, StringComparison.OrdinalIgnoreCase))
                    return i;
            return 0;
        }

        private void OnSaveClick(object sender, EventArgs e)
        {
            _settings.OverridePerModel = _overrideCheck.Checked;
            _settings.ModelPreviewCount = AiImageService.ClampModelPreviewCount((int)_previewCountNum.Value);

            if (int.TryParse((_widthCombo.Text ?? string.Empty).Trim(), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int w) && w > 0)
                _settings.DefaultWidth = Math.Max(1, Math.Min(4096, w));

            if (int.TryParse((_heightCombo.Text ?? string.Empty).Trim(), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int h) && h > 0)
                _settings.DefaultHeight = Math.Max(1, Math.Min(4096, h));

            _settings.DefaultQuality = QualityValues[_qualityCombo.SelectedIndex];
            _settings.DefaultFormat = FormatValues[_formatCombo.SelectedIndex];

            DialogResult = DialogResult.OK;
            Close();
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
    }
}
