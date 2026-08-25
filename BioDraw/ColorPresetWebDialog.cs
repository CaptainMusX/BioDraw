using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class ColorPresetWebDialog : WebViewDialogBase
    {
        private readonly string _title;
        private readonly List<string> _options;
        private readonly string _currentValue;
        private readonly bool _allowEmpty;
        private readonly Action _onPersist;
        private readonly Action _onInvalidate;
        private readonly Func<string, string> _onPickColor;
        private string _selectedValue;

        public string SelectedValue
        {
            get { return _selectedValue; }
        }

        public ColorPresetWebDialog(
            string title,
            List<string> options,
            string currentValue,
            bool allowEmpty,
            Func<string, string> onPickColor,
            Action onPersist,
            Action onInvalidate)
        {
            _title = title;
            _options = options ?? new List<string>();
            _currentValue = PresetManager.NormalizeColorInputText(currentValue);
            _allowEmpty = allowEmpty;
            _selectedValue = _currentValue;
            _onPickColor = onPickColor;
            _onPersist = onPersist;
            _onInvalidate = onInvalidate;
            Text = title;
            InitWebView(600, 420, "color-preset.html");
        }

        protected override void OnWebViewReady()
        {
            PostMessageToWeb("init", new Dictionary<string, object>
            {
                { "title", _title ?? string.Empty },
                { "currentValue", _currentValue },
                { "allowEmpty", _allowEmpty },
                { "colors", _options.ToArray() }
            });
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
                case "pickColor":
                    PickColor(payload);
                    break;
            }
        }

        private void ApplySave(Dictionary<string, object> payload)
        {
            string value;
            if (!PresetManager.TryNormalizeImageMagickColor(
                    GetString(payload, "value"), _allowEmpty, out value))
            {
                PostMessageToWeb("validationError", new Dictionary<string, object>
                {
                    { "message", "颜色格式无效。请使用颜色名称、十六进制颜色或 rgb/rgba 表达式。" }
                });
                return;
            }

            SyncOptions(payload);

            if (!string.IsNullOrEmpty(value))
            {
                var position = ClampInt(
                    GetInt(payload, "position", _options.Count + 1),
                    1,
                    Math.Max(1, _options.Count + 1));
                PresetManager.UpsertColorOptionAtPosition(_options, value, position);
            }

            _selectedValue = value;
            if (_onPersist != null) _onPersist();
            if (_onInvalidate != null) _onInvalidate();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void SyncOptions(Dictionary<string, object> payload)
        {
            object colorsValue;
            var colors = payload.TryGetValue("colors", out colorsValue)
                ? colorsValue as object[]
                : null;
            if (colors == null) return;

            _options.Clear();
            foreach (var item in colors)
            {
                string normalized;
                if (!PresetManager.TryNormalizeImageMagickColor(
                        item as string, false, out normalized))
                    continue;
                PresetManager.AddColorOption(_options, normalized);
            }
        }

        private void PickColor(Dictionary<string, object> payload)
        {
            if (_onPickColor == null) return;
            var initial = PresetManager.NormalizeColorInputText(
                GetString(payload, "value"));
            if (string.IsNullOrWhiteSpace(initial)) initial = _selectedValue;
            var picked = _onPickColor(initial);
            if (string.IsNullOrWhiteSpace(picked)) return;
            PostMessageToWeb("colorPicked", new Dictionary<string, object>
            {
                { "value", picked }
            });
        }
    }
}
