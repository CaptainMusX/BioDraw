using System;
using System.Collections.Generic;
using System.Reflection;

namespace BioDraw
{
    internal sealed class AboutDialog : WebViewDialogBase
    {
        public AboutDialog()
        {
            Text = "关于 BioDraw";
            InitWebView(420, 380, "about.html");
        }

        protected override void OnWebViewReady()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            PostMessageToWeb("init", new Dictionary<string, object>
            {
                { "version", version == null ? "" : version.ToString(3) }
            });
        }

        protected override void HandleWebMessage(
            string action, Dictionary<string, object> payload)
        {
            if (string.Equals(action, "cancel", StringComparison.Ordinal)) Close();
        }
    }
}
