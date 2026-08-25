using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class PresetEditorWebDialog : WebViewDialogBase
    {
        private readonly ImageReplacePreset _source;
        private readonly bool _canDelete;
        private readonly int _maximumSortOrder;
        private readonly bool _initialSetAsDefault;
        private ImageReplacePreset _result;
        private bool _setAsDefault;
        private bool _deleteRequested;

        public ImageReplacePreset Result { get { return _result; } }
        public bool SetAsDefault { get { return _setAsDefault; } }
        public bool DeleteRequested { get { return _deleteRequested; } }

        public PresetEditorWebDialog(
            ImageReplacePreset source,
            bool canDelete,
            int maximumSortOrder,
            bool initialSetAsDefault)
        {
            _source = source ?? PresetManager.CreateDefaultPreset();
            _canDelete = canDelete;
            _maximumSortOrder = Math.Max(1, maximumSortOrder);
            _initialSetAsDefault = initialSetAsDefault;
            Text = "颜色替换参数";
            InitWebView(520, 380, "preset-editor.html");
        }

        protected override void OnWebViewReady()
        {
            PostMessageToWeb("init", new Dictionary<string, object>
            {
                { "name", _source.Name ?? string.Empty },
                { "sortOrder", ClampInt(_source.SortOrder, 1, _maximumSortOrder) },
                { "maximumSortOrder", _maximumSortOrder },
                { "fuzzPercent", Math.Round(PresetManager.NormalizeFuzzPercent(_source.FuzzPercent), 1) },
                { "canDelete", _canDelete },
                { "setAsDefault", _initialSetAsDefault }
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
                case "delete":
                    ConfirmDelete();
                    break;
            }
        }

        private void ApplySave(Dictionary<string, object> payload)
        {
            var name = GetString(payload, "name").Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                PostMessageToWeb("validationError", new Dictionary<string, object>
                {
                    { "message", "名称不能为空。" }
                });
                return;
            }
            if (name.Length > 100) name = name.Substring(0, 100);

            _result = new ImageReplacePreset
            {
                Name = name,
                SortOrder = ClampInt(
                    GetInt(payload, "sortOrder", 1), 1, _maximumSortOrder),
                TargetColor = _source.TargetColor,
                Mode = _source.Mode,
                ReplacementColor = _source.ReplacementColor,
                FuzzPercent = PresetManager.NormalizeFuzzPercent(
                    GetDouble(payload, "fuzzPercent", 0))
            };
            _setAsDefault = GetBool(payload, "setAsDefault");
            DialogResult = DialogResult.OK;
            Close();
        }

        private void ConfirmDelete()
        {
            if (!_canDelete) return;
            var choice = MessageBox.Show(
                "确认删除预设 \"" + (_source.Name ?? string.Empty) + "\" 吗？",
                "BioDraw",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (choice != DialogResult.Yes) return;

            _deleteRequested = true;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
