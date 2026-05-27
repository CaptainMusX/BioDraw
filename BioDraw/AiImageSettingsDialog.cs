using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class AiImageSettingsDialog : Form
    {
        private readonly AiImageApiSettings _settings;
        private readonly List<AiImageApiSettings> _allSettings;
        private TextBox _displayNameBox;
        private TextBox _apiTokenBox;
        private TextBox _endpointBox;
        private TextBox _modelBox;
        private ComboBox _widthCombo;
        private ComboBox _heightCombo;
        private TrackBar _widthSlider;
        private TrackBar _heightSlider;
        private ComboBox _qualityCombo;
        private ComboBox _formatCombo;
        private PictureBox _iconPictureBox;
        private Button _lockAspectButton;
        private bool _lockAspect;
        private bool _updatingSizeControls;

        private static readonly string[] AiSizeOptions =
        {
            "512", "768", "1024", "1536", "1792", "1920",
            "2048", "2160", "2560", "3072", "3840", "4096"
        };

        private static readonly string[] QualityOptions = { "auto (自动)", "low (低)", "medium (中)", "high (高)" };
        private static readonly string[] QualityValues = { "auto", "low", "medium", "high" };
        private static readonly string[] FormatOptions = { "PNG", "JPEG", "WebP" };
        private static readonly string[] FormatValues = { "png", "jpeg", "webp" };

        public AiImageSettingsDialog(AiImageApiSettings settings, List<AiImageApiSettings> allSettings)
        {
            _settings = settings ?? AiImageService.CreateDefaultSettings();
            _allSettings = allSettings ?? new List<AiImageApiSettings> { _settings };
            _lockAspect = _settings.LockAspectRatio;
            InitializeForm();
        }

        private void InitializeForm()
        {
            Text = "模型设置 - " + (_settings.DisplayName ?? "未命名");
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Regular, GraphicsUnit.Point);
            BackColor = Color.FromArgb(244, 247, 252);
            ForeColor = Color.FromArgb(32, 41, 57);
            AutoScaleMode = AutoScaleMode.Dpi;
            MaximizeBox = true;
            MinimizeBox = false;
            MinimumSize = new Size(640, 580);
            ClientSize = new Size(680, 620);
            Icon = SystemIcons.Application;

            // ---- Icon (click to choose) ----
            _iconPictureBox = new PictureBox
            {
                Size = new Size(56, 56),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Cursor = Cursors.Hand
            };
            _iconPictureBox.Click += (s, e) => ChooseIcon();
            LoadIconPreview();

            // Label widths must fit "名称"/"Key"/"地址"/"模型"/"宽度"/"高度"/"质量"/"格式"
            // At 11pt "Microsoft YaHei UI", 2 Chinese chars ≈ 28px; add padding for right-align → 56px min.
            const int labelWidth = 80;
            const int fieldGap = 10;
            const int rowHeight = 32;
            const int rowGap = 14;
            const int margin = 20;

            var lblDisplayName = new Label { Text = "名称", TextAlign = ContentAlignment.MiddleRight };
            _displayNameBox = CreateTextBox();
            _displayNameBox.Text = _settings.DisplayName ?? string.Empty;

            var lblApiToken = new Label { Text = "Key", TextAlign = ContentAlignment.MiddleRight };
            _apiTokenBox = CreateTextBox();
            _apiTokenBox.Text = _settings.ApiToken ?? string.Empty;
            _apiTokenBox.UseSystemPasswordChar = true;

            var showTokenCheck = new CheckBox
            {
                Text = "显示",
                AutoSize = true,
                ForeColor = Color.FromArgb(43, 52, 69)
            };
            showTokenCheck.CheckedChanged += (s, e) =>
                _apiTokenBox.UseSystemPasswordChar = !showTokenCheck.Checked;

            var lblEndpoint = new Label { Text = "地址", TextAlign = ContentAlignment.MiddleRight };
            _endpointBox = CreateTextBox();
            _endpointBox.Text = _settings.EndpointUrl ?? string.Empty;

            var lblModel = new Label { Text = "模型", TextAlign = ContentAlignment.MiddleRight };
            _modelBox = CreateTextBox();
            _modelBox.Text = _settings.Model ?? string.Empty;

            var lblWidth = new Label { Text = "宽度", TextAlign = ContentAlignment.MiddleRight };
            _widthCombo = CreateEditableCombo(AiSizeOptions, GetClosestSizeIndex(_settings.DefaultWidth));
            _widthCombo.Text = _settings.DefaultWidth.ToString(CultureInfo.InvariantCulture);
            _widthCombo.TextChanged += OnWidthComboChanged;
            _widthSlider = CreateTrackBar();
            _widthSlider.Value = ClampSliderValue(_settings.DefaultWidth);
            _widthSlider.ValueChanged += OnWidthSliderChanged;

            var lblHeight = new Label { Text = "高度", TextAlign = ContentAlignment.MiddleRight };
            _heightCombo = CreateEditableCombo(AiSizeOptions, GetClosestSizeIndex(_settings.DefaultHeight));
            _heightCombo.Text = _settings.DefaultHeight.ToString(CultureInfo.InvariantCulture);
            _heightCombo.TextChanged += OnHeightComboChanged;
            _heightSlider = CreateTrackBar();
            _heightSlider.Value = ClampSliderValue(_settings.DefaultHeight);
            _heightSlider.ValueChanged += OnHeightSliderChanged;

            // Lock aspect button — centered vertically across both size rows
            _lockAspectButton = new Button
            {
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand,
                Size = new Size(32, 32),
                Font = new Font("Segoe UI Emoji", 13F),
                Text = _lockAspect ? "🔗" : "⬚"
            };
            _lockAspectButton.FlatAppearance.BorderSize = 1;
            _lockAspectButton.FlatAppearance.BorderColor = _lockAspect
                ? Color.FromArgb(24, 118, 242) : Color.FromArgb(189, 198, 213);
            _lockAspectButton.BackColor = _lockAspect
                ? Color.FromArgb(225, 240, 255) : Color.White;
            _lockAspectButton.Click += (s, e) =>
            {
                _lockAspect = !_lockAspect;
                _lockAspectButton.Text = _lockAspect ? "🔗" : "⬚";
                _lockAspectButton.FlatAppearance.BorderColor = _lockAspect
                    ? Color.FromArgb(24, 118, 242) : Color.FromArgb(189, 198, 213);
                _lockAspectButton.BackColor = _lockAspect
                    ? Color.FromArgb(225, 240, 255) : Color.White;
            };

            var lblQuality = new Label { Text = "质量", TextAlign = ContentAlignment.MiddleRight };
            _qualityCombo = CreateCombo(QualityOptions, GetQualityIndex(_settings.DefaultQuality));

            var lblFormat = new Label { Text = "格式", TextAlign = ContentAlignment.MiddleRight };
            _formatCombo = CreateCombo(FormatOptions, GetFormatIndex(_settings.DefaultFormat));

            var btnDelete = new Button
            {
                Text = "删除",
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand
            };
            btnDelete.FlatAppearance.BorderSize = 1;
            btnDelete.FlatAppearance.BorderColor = Color.FromArgb(220, 53, 69);
            btnDelete.BackColor = Color.FromArgb(220, 53, 69);
            btnDelete.ForeColor = Color.White;
            btnDelete.Click += OnDeleteClick;

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

            Controls.Add(_iconPictureBox);
            Controls.Add(lblDisplayName);
            Controls.Add(_displayNameBox);
            Controls.Add(lblApiToken);
            Controls.Add(_apiTokenBox);
            Controls.Add(showTokenCheck);
            Controls.Add(lblEndpoint);
            Controls.Add(_endpointBox);
            Controls.Add(lblModel);
            Controls.Add(_modelBox);
            Controls.Add(lblWidth);
            Controls.Add(_widthCombo);
            Controls.Add(_widthSlider);
            Controls.Add(lblHeight);
            Controls.Add(_heightCombo);
            Controls.Add(_heightSlider);
            Controls.Add(_lockAspectButton);
            Controls.Add(lblQuality);
            Controls.Add(_qualityCombo);
            Controls.Add(lblFormat);
            Controls.Add(_formatCombo);
            Controls.Add(btnDelete);
            Controls.Add(btnSave);
            Controls.Add(btnCancel);

            AcceptButton = btnSave;

            Resize += (_, __) => LayoutControls();
            LayoutControls();

            void LayoutControls()
            {
                var y = margin;
                var fieldX = margin + labelWidth + fieldGap;
                // Sliders get ~half the remaining space, combos get fixed width
                var comboWidth = 110;
                var sliderWidth = Math.Max(180, ClientSize.Width - fieldX - comboWidth - 48 - margin);
                var fullWidth = ClientSize.Width - fieldX - margin;

                // Icon centered
                _iconPictureBox.Location = new Point((ClientSize.Width - 56) / 2, y);
                y += 56 + rowGap + 4;

                LayoutRow(lblDisplayName, _displayNameBox, fullWidth, labelWidth, fieldX, ref y);
                LayoutRow(lblApiToken, _apiTokenBox, fullWidth, labelWidth, fieldX, ref y);
                showTokenCheck.Location = new Point(fieldX, y);
                y += rowHeight + rowGap;

                LayoutRow(lblEndpoint, _endpointBox, fullWidth, labelWidth, fieldX, ref y);
                LayoutRow(lblModel, _modelBox, fullWidth, labelWidth, fieldX, ref y);

                // Width row (lock button spans width+height rows, positioned after both are placed)
                int widthRowY = y;
                lblWidth.Location = new Point(margin, widthRowY);
                lblWidth.Size = new Size(labelWidth, rowHeight);
                _widthCombo.Location = new Point(fieldX, widthRowY);
                _widthCombo.Size = new Size(comboWidth, rowHeight);
                _widthSlider.Location = new Point(fieldX + comboWidth + 8, widthRowY);
                _widthSlider.Size = new Size(sliderWidth, rowHeight);
                y += rowHeight + rowGap;

                // Height row
                int heightRowY = y;
                lblHeight.Location = new Point(margin, heightRowY);
                lblHeight.Size = new Size(labelWidth, rowHeight);
                _heightCombo.Location = new Point(fieldX, heightRowY);
                _heightCombo.Size = new Size(comboWidth, rowHeight);
                _heightSlider.Location = new Point(fieldX + comboWidth + 8, heightRowY);
                _heightSlider.Size = new Size(sliderWidth, rowHeight);
                y += rowHeight + rowGap;

                // Lock button: vertically centered across both size rows
                int lockX = fieldX + comboWidth + 8 + sliderWidth + 8;
                int lockY = widthRowY + (heightRowY + rowHeight - widthRowY - 32) / 2;
                _lockAspectButton.Location = new Point(lockX, lockY);

                LayoutRow(lblQuality, _qualityCombo, Math.Max(comboWidth, Math.Min(200, fullWidth)), labelWidth, fieldX, ref y);
                LayoutRow(lblFormat, _formatCombo, Math.Max(comboWidth, Math.Min(200, fullWidth)), labelWidth, fieldX, ref y);

                var bottomY = ClientSize.Height - margin - 40;
                const int buttonWidth = 100;
                const int buttonHeight = 40;
                var rightEdge = ClientSize.Width - margin;

                btnDelete.Location = new Point(margin, bottomY);
                btnDelete.Size = new Size(buttonWidth, buttonHeight);
                btnSave.Location = new Point(rightEdge - buttonWidth, bottomY);
                btnSave.Size = new Size(buttonWidth, buttonHeight);
                btnCancel.Location = new Point(rightEdge - buttonWidth * 2 - 12, bottomY);
                btnCancel.Size = new Size(buttonWidth, buttonHeight);
            }

            void LayoutRow(Control label, Control field, int fieldWidth, int lw, int fx, ref int y2)
            {
                label.Location = new Point(margin, y2);
                label.Size = new Size(lw, rowHeight);
                field.Location = new Point(fx, y2);
                field.Size = new Size(fieldWidth, rowHeight);
                y2 += rowHeight + rowGap;
            }
        }

        private void LoadIconPreview()
        {
            if (!string.IsNullOrWhiteSpace(_settings.IconPath) && File.Exists(_settings.IconPath))
            {
                try
                {
                    using (var img = Image.FromFile(_settings.IconPath))
                        _iconPictureBox.Image = new Bitmap(img);
                    return;
                }
                catch { }
            }
            // Default placeholder
            var bmp = new Bitmap(56, 56);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.FromArgb(230, 235, 245));
                using (var pen = new Pen(Color.FromArgb(160, 168, 185), 2f))
                {
                    g.DrawLine(pen, 20, 28, 36, 28);
                    g.DrawLine(pen, 28, 20, 28, 36);
                }
            }
            _iconPictureBox.Image = bmp;
        }

        private void ChooseIcon()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "选择图标";
                dialog.Filter = "图像文件|*.png;*.ico;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff";
                dialog.CheckFileExists = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _settings.IconPath = dialog.FileName;
                    LoadIconPreview();
                }
            }
        }

        private static TrackBar CreateTrackBar()
        {
            return new TrackBar
            {
                Minimum = 1,
                Maximum = 4096,
                TickFrequency = 512,
                SmallChange = 32,
                LargeChange = 256,
                AutoSize = false
            };
        }

        private static int ClampSliderValue(int value)
        {
            return Math.Max(1, Math.Min(4096, value));
        }

        private void SyncWidthFromCombo()
        {
            if (_updatingSizeControls) return;
            if (!int.TryParse((_widthCombo.Text ?? string.Empty).Trim(), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int w) || w <= 0)
                return;
            _updatingSizeControls = true;
            w = Math.Max(1, Math.Min(4096, w));
            _widthSlider.Value = ClampSliderValue(w);
            if (_lockAspect && _settings.DefaultWidth > 0)
            {
                int h = (int)Math.Round((double)w * _settings.DefaultHeight / _settings.DefaultWidth);
                h = Math.Max(1, Math.Min(4096, h));
                _heightCombo.Text = h.ToString(CultureInfo.InvariantCulture);
                _heightSlider.Value = ClampSliderValue(h);
            }
            _updatingSizeControls = false;
        }

        private void SyncHeightFromCombo()
        {
            if (_updatingSizeControls) return;
            if (!int.TryParse((_heightCombo.Text ?? string.Empty).Trim(), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int h) || h <= 0)
                return;
            _updatingSizeControls = true;
            h = Math.Max(1, Math.Min(4096, h));
            _heightSlider.Value = ClampSliderValue(h);
            if (_lockAspect && _settings.DefaultHeight > 0)
            {
                int w = (int)Math.Round((double)h * _settings.DefaultWidth / _settings.DefaultHeight);
                w = Math.Max(1, Math.Min(4096, w));
                _widthCombo.Text = w.ToString(CultureInfo.InvariantCulture);
                _widthSlider.Value = ClampSliderValue(w);
            }
            _updatingSizeControls = false;
        }

        private void OnWidthComboChanged(object sender, EventArgs e) => SyncWidthFromCombo();

        private void OnWidthSliderChanged(object sender, EventArgs e)
        {
            if (_updatingSizeControls) return;
            _updatingSizeControls = true;
            int w = _widthSlider.Value;
            _widthCombo.Text = w.ToString(CultureInfo.InvariantCulture);
            if (_lockAspect && _settings.DefaultWidth > 0)
            {
                int h = (int)Math.Round((double)w * _settings.DefaultHeight / _settings.DefaultWidth);
                h = Math.Max(1, Math.Min(4096, h));
                _heightCombo.Text = h.ToString(CultureInfo.InvariantCulture);
                _heightSlider.Value = ClampSliderValue(h);
            }
            _updatingSizeControls = false;
        }

        private void OnHeightComboChanged(object sender, EventArgs e) => SyncHeightFromCombo();

        private void OnHeightSliderChanged(object sender, EventArgs e)
        {
            if (_updatingSizeControls) return;
            _updatingSizeControls = true;
            int h = _heightSlider.Value;
            _heightCombo.Text = h.ToString(CultureInfo.InvariantCulture);
            if (_lockAspect && _settings.DefaultHeight > 0)
            {
                int w = (int)Math.Round((double)h * _settings.DefaultWidth / _settings.DefaultHeight);
                w = Math.Max(1, Math.Min(4096, w));
                _widthCombo.Text = w.ToString(CultureInfo.InvariantCulture);
                _widthSlider.Value = ClampSliderValue(w);
            }
            _updatingSizeControls = false;
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
            var token = (_apiTokenBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                var result = MessageBox.Show(
                    "API Key 为空，将无法正常生图。\n\n是否继续保存？",
                    "BioDraw",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result != DialogResult.Yes)
                    return;
            }

            _settings.DisplayName = (_displayNameBox.Text ?? string.Empty).Trim();
            _settings.ApiToken = token;
            _settings.EndpointUrl = (_endpointBox.Text ?? string.Empty).Trim();
            _settings.Model = (_modelBox.Text ?? string.Empty).Trim();
            _settings.DefaultQuality = QualityValues[_qualityCombo.SelectedIndex];
            _settings.DefaultFormat = FormatValues[_formatCombo.SelectedIndex];
            _settings.LockAspectRatio = _lockAspect;

            if (int.TryParse((_widthCombo.Text ?? string.Empty).Trim(), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int w) && w > 0)
                _settings.DefaultWidth = Math.Max(1, Math.Min(4096, w));

            if (int.TryParse((_heightCombo.Text ?? string.Empty).Trim(), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int h) && h > 0)
                _settings.DefaultHeight = Math.Max(1, Math.Min(4096, h));

            // Add to list if not already in it
            if (!_allSettings.Contains(_settings))
                _allSettings.Add(_settings);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void OnDeleteClick(object sender, EventArgs e)
        {
            if (_allSettings.Count <= 1)
            {
                MessageBox.Show("至少需要保留 1 个模型配置。", "BioDraw",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                "确认删除模型 \"" + (_settings.DisplayName ?? _settings.Model) + "\" 吗？\n\n此操作不可撤销。",
                "BioDraw",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
                return;

            _allSettings.Remove(_settings);

            var globalSettings = AiImageService.LoadGlobalSettings();
            AiImageService.SaveSettings(_allSettings, globalSettings);

            DialogResult = DialogResult.OK;
            Close();
        }

        private static TextBox CreateTextBox()
        {
            return new TextBox
            {
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(32, 41, 57)
            };
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
