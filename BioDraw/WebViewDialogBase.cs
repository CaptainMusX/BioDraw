using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace BioDraw
{
    internal class WebViewDialogBase : Form
    {
        protected WebView2 webView;

        private static readonly Dictionary<string, Rectangle> SavedBounds =
            new Dictionary<string, Rectangle>();
        private static readonly object EnvironmentSync = new object();
        private static readonly JavaScriptSerializer JsonSerializer = new JavaScriptSerializer();
        private static Task<CoreWebView2Environment> _environmentTask;

        private bool _webViewReady;
        private string _boundsKey;
        private string _htmlResourceName;

        private const int TitleBarHeight = 44;
        private const int TrafficLightsWidth = 96;
        private const int ResizeGrip = 8;
        private const int MaxWebMessageLength = 256 * 1024;

        protected WebViewDialogBase()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.FromArgb(235, 238, 245);
            KeyPreview = true;
        }

        protected void InitWebView(int width, int height, string htmlResourceName, string boundsKey = null)
        {
            ClientSize = new Size(width, height);
            MinimumSize = new Size(Math.Min(360, width), Math.Min(240, height));

            _boundsKey = boundsKey ?? GetType().Name;
            _htmlResourceName = htmlResourceName;

            var saved = LoadSavedBounds(_boundsKey);
            if (saved.HasValue)
            {
                StartPosition = FormStartPosition.Manual;
                Bounds = EnsureVisibleBounds(saved.Value);
            }

            FormClosed += (s, e) => PersistBounds(_boundsKey, Bounds);

            webView = new WebView2 { Dock = DockStyle.Fill };
            webView.NavigationCompleted += OnNavigationCompleted;
            webView.WebMessageReceived += OnWebMessageReceived;
            Controls.Add(webView);

            AddWindowInteractionOverlays(width);
            Shown += OnDialogShown;
        }

        private void AddWindowInteractionOverlays(int initialWidth)
        {
            var dragOverlay = new Panel
            {
                BackColor = Color.Transparent,
                Cursor = Cursors.SizeAll,
                Left = TrafficLightsWidth,
                Top = 0,
                Height = TitleBarHeight,
                Width = Math.Max(1, initialWidth - TrafficLightsWidth)
            };
            dragOverlay.MouseDown += (s, e) => BeginNativeMove(e);
            Resize += (s, e) =>
                dragOverlay.Width = Math.Max(1, ClientSize.Width - TrafficLightsWidth);
            Controls.Add(dragOverlay);
            dragOverlay.BringToFront();

            AddResizeGrip(Cursors.SizeNWSE,
                () => new Rectangle(0, 0, ResizeGrip, ResizeGrip), 13);
            AddResizeGrip(Cursors.SizeNS,
                () => new Rectangle(ResizeGrip, 0, Math.Max(1, ClientSize.Width - ResizeGrip * 2), 6), 12);
            AddResizeGrip(Cursors.SizeNESW,
                () => new Rectangle(ClientSize.Width - ResizeGrip, 0, ResizeGrip, ResizeGrip), 14);
            AddResizeGrip(Cursors.SizeWE,
                () => new Rectangle(0, ResizeGrip, 6, Math.Max(1, ClientSize.Height - ResizeGrip * 2)), 10);
            AddResizeGrip(Cursors.SizeWE,
                () => new Rectangle(ClientSize.Width - 6, ResizeGrip, 6, Math.Max(1, ClientSize.Height - ResizeGrip * 2)), 11);
            AddResizeGrip(Cursors.SizeNESW,
                () => new Rectangle(0, ClientSize.Height - ResizeGrip, ResizeGrip, ResizeGrip), 16);
            AddResizeGrip(Cursors.SizeNS,
                () => new Rectangle(ResizeGrip, ClientSize.Height - 6, Math.Max(1, ClientSize.Width - ResizeGrip * 2), 6), 15);
            AddResizeGrip(Cursors.SizeNWSE,
                () => new Rectangle(ClientSize.Width - ResizeGrip, ClientSize.Height - ResizeGrip, ResizeGrip, ResizeGrip), 17);
        }

        private void BeginNativeMove(MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            NativeMethods.ReleaseCapture();
            NativeMethods.SendMessage(Handle, NativeMethods.WM_NCLBUTTONDOWN,
                NativeMethods.HT_CAPTION, 0);
        }

        private void AddResizeGrip(Cursor cursor, Func<Rectangle> getBounds, int htCode)
        {
            var grip = new Panel
            {
                BackColor = Color.Transparent,
                Cursor = cursor,
                Bounds = getBounds()
            };
            grip.MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Left) return;
                NativeMethods.ReleaseCapture();
                NativeMethods.SendMessage(Handle, NativeMethods.WM_NCLBUTTONDOWN, htCode, 0);
            };
            Resize += (s, e) => grip.Bounds = getBounds();
            Controls.Add(grip);
            grip.BringToFront();
        }

        private async void OnDialogShown(object sender, EventArgs e)
        {
            ApplyDwmRoundCorners();
            try
            {
                var environment = await GetSharedEnvironmentAsync();
                if (IsDisposed || Disposing) return;

                await webView.EnsureCoreWebView2Async(environment);
                if (IsDisposed || Disposing) return;

                ConfigureWebView(webView.CoreWebView2);
                webView.CoreWebView2.NavigateToString(LoadEmbeddedHtml(_htmlResourceName));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("BioDraw WebView2 initialization failed: " + ex);
                if (!IsDisposed && !Disposing)
                {
                    MessageBox.Show(
                        "界面初始化失败。请确认已安装 Microsoft Edge WebView2 Runtime，然后重试。\n\n" + ex.Message,
                        "BioDraw",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    Close();
                }
            }
        }

        private static Task<CoreWebView2Environment> GetSharedEnvironmentAsync()
        {
            lock (EnvironmentSync)
            {
                if (_environmentTask == null || _environmentTask.IsFaulted || _environmentTask.IsCanceled)
                {
                    var folder = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "BioDraw", "WebView2");
                    Directory.CreateDirectory(folder);
                    _environmentTask = CoreWebView2Environment.CreateAsync(null, folder);
                }
                return _environmentTask;
            }
        }

        private static void ConfigureWebView(CoreWebView2 core)
        {
            core.Settings.AreDevToolsEnabled = false;
            core.Settings.AreDefaultContextMenusEnabled = false;
            core.Settings.AreDefaultScriptDialogsEnabled = false;
            core.Settings.AreHostObjectsAllowed = false;
            core.Settings.IsStatusBarEnabled = false;
            core.Settings.IsZoomControlEnabled = true;
            core.Settings.IsGeneralAutofillEnabled = false;
            core.Settings.IsPasswordAutosaveEnabled = false;

            core.NavigationStarting += (s, e) =>
            {
                if (!string.Equals(e.Uri, "about:blank", StringComparison.OrdinalIgnoreCase))
                    e.Cancel = true;
            };
            core.NewWindowRequested += (s, e) => e.Handled = true;
            core.DownloadStarting += (s, e) =>
            {
                e.Cancel = true;
                e.Handled = true;
            };
            core.PermissionRequested += (s, e) =>
            {
                e.State = CoreWebView2PermissionState.Deny;
                e.Handled = true;
            };
        }

        private void ApplyDwmRoundCorners()
        {
            try
            {
                int value = NativeMethods.DWMWCP_ROUND;
                NativeMethods.DwmSetWindowAttribute(
                    Handle, NativeMethods.DWMWA_WINDOW_CORNER_PREFERENCE,
                    ref value, sizeof(int));
            }
            catch
            {
                // Cosmetic only; older Windows versions do not expose this attribute.
            }
        }

        private void OnNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!e.IsSuccess) return;
            _webViewReady = true;
            OnWebViewReady();
        }

        protected virtual void OnWebViewReady()
        {
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var raw = e.WebMessageAsJson;
                if (string.IsNullOrWhiteSpace(raw) || raw.Length > MaxWebMessageLength) return;

                object decoded = JsonSerializer.DeserializeObject(raw);
                var encodedString = decoded as string;
                if (encodedString != null)
                {
                    if (encodedString.Length > MaxWebMessageLength) return;
                    decoded = JsonSerializer.DeserializeObject(encodedString);
                }

                var message = decoded as Dictionary<string, object>;
                if (message == null) return;

                object actionValue;
                var action = message.TryGetValue("action", out actionValue)
                    ? actionValue as string
                    : null;
                if (string.IsNullOrWhiteSpace(action)) return;

                object payloadValue;
                var payload = message.TryGetValue("payload", out payloadValue)
                    ? payloadValue as Dictionary<string, object>
                    : null;
                payload = payload ?? new Dictionary<string, object>();

                if (string.Equals(action, "drag", StringComparison.Ordinal))
                {
                    NativeMethods.ReleaseCapture();
                    NativeMethods.SendMessage(Handle, NativeMethods.WM_NCLBUTTONDOWN,
                        NativeMethods.HT_CAPTION, 0);
                    return;
                }

                if (string.Equals(action, "resize", StringComparison.Ordinal))
                {
                    int ht = ParseResizeHt(payload);
                    if (ht > 0)
                    {
                        NativeMethods.ReleaseCapture();
                        NativeMethods.SendMessage(Handle, NativeMethods.WM_NCLBUTTONDOWN, ht, 0);
                    }
                    return;
                }

                HandleWebMessage(action, payload);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("BioDraw ignored an invalid WebView message: " + ex.Message);
            }
        }

        private static int ParseResizeHt(Dictionary<string, object> payload)
        {
            object directionValue;
            var direction = payload.TryGetValue("direction", out directionValue)
                ? directionValue as string
                : null;
            switch (direction)
            {
                case "nw": return 13;
                case "n": return 12;
                case "ne": return 14;
                case "w": return 10;
                case "e": return 11;
                case "sw": return 16;
                case "s": return 15;
                case "se": return 17;
                default: return 0;
            }
        }

        protected virtual void HandleWebMessage(
            string action, Dictionary<string, object> payload)
        {
        }

        protected void PostMessageToWeb(string action, object payload = null)
        {
            if (!_webViewReady || webView == null || webView.CoreWebView2 == null) return;
            var message = new Dictionary<string, object>
            {
                { "action", action },
                { "payload", payload ?? new Dictionary<string, object>() }
            };
            webView.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(message));
        }

        protected void ExecuteScript(string script)
        {
            if (!_webViewReady || webView == null || webView.CoreWebView2 == null) return;
            webView.CoreWebView2.ExecuteScriptAsync(script);
        }

        protected void TryBeginInvoke(Action action)
        {
            if (action == null || IsDisposed || Disposing || !IsHandleCreated) return;
            try
            {
                BeginInvoke(new Action(() =>
                {
                    if (!IsDisposed && !Disposing) action();
                }));
            }
            catch (ObjectDisposedException)
            {
            }
            catch (InvalidOperationException)
            {
            }
        }

        protected void CloseAfter(int milliseconds)
        {
            var timer = new Timer { Interval = Math.Max(1, milliseconds) };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                if (!IsDisposed && !Disposing) Close();
            };
            timer.Start();
        }

        private static string BoundsFilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "BioDraw", "WebDialogBounds.xml");
            }
        }

        private static Rectangle? LoadSavedBounds(string key)
        {
            Rectangle cached;
            if (SavedBounds.TryGetValue(key, out cached)) return cached;
            try
            {
                if (!File.Exists(BoundsFilePath)) return null;
                var doc = System.Xml.Linq.XDocument.Load(BoundsFilePath);
                var element = doc.Root == null ? null : doc.Root.Element(SafeXmlName(key));
                if (element == null) return null;
                int x = (int?)element.Attribute("X") ?? 0;
                int y = (int?)element.Attribute("Y") ?? 0;
                int w = (int?)element.Attribute("W") ?? 0;
                int h = (int?)element.Attribute("H") ?? 0;
                if (w < 100 || h < 60) return null;
                return EnsureVisibleBounds(new Rectangle(x, y, w, h));
            }
            catch
            {
                return null;
            }
        }

        private static Rectangle EnsureVisibleBounds(Rectangle requested)
        {
            Screen target = null;
            int bestArea = 0;
            foreach (var screen in Screen.AllScreens)
            {
                var intersection = Rectangle.Intersect(screen.WorkingArea, requested);
                int area = Math.Max(0, intersection.Width) * Math.Max(0, intersection.Height);
                if (area <= bestArea) continue;
                bestArea = area;
                target = screen;
            }

            target = target ?? Screen.PrimaryScreen;
            if (target == null) return requested;

            var work = target.WorkingArea;
            int width = Math.Min(Math.Max(100, requested.Width), work.Width);
            int height = Math.Min(Math.Max(60, requested.Height), work.Height);
            int x = Math.Max(work.Left, Math.Min(requested.X, work.Right - width));
            int y = Math.Max(work.Top, Math.Min(requested.Y, work.Bottom - height));
            return new Rectangle(x, y, width, height);
        }

        private static void PersistBounds(string key, Rectangle bounds)
        {
            bounds = EnsureVisibleBounds(bounds);
            SavedBounds[key] = bounds;
            try
            {
                var path = BoundsFilePath;
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);

                System.Xml.Linq.XDocument doc;
                try
                {
                    doc = File.Exists(path) ? System.Xml.Linq.XDocument.Load(path) : null;
                }
                catch
                {
                    doc = null;
                }

                if (doc == null || doc.Root == null)
                    doc = new System.Xml.Linq.XDocument(
                        new System.Xml.Linq.XElement("WebDialogBounds"));

                var xname = SafeXmlName(key);
                var element = doc.Root.Element(xname);
                if (element == null)
                {
                    element = new System.Xml.Linq.XElement(xname);
                    doc.Root.Add(element);
                }
                element.SetAttributeValue("X", bounds.X);
                element.SetAttributeValue("Y", bounds.Y);
                element.SetAttributeValue("W", bounds.Width);
                element.SetAttributeValue("H", bounds.Height);

                var tempPath = path + ".tmp";
                doc.Save(tempPath);
                if (File.Exists(path))
                    File.Replace(tempPath, path, null);
                else
                    File.Move(tempPath, path);
            }
            catch
            {
                // Bounds persistence must never prevent a dialog from closing.
            }
        }

        private static string SafeXmlName(string key)
        {
            var sb = new StringBuilder("D_");
            foreach (char c in key)
                sb.Append(char.IsLetterOrDigit(c) ? c : '_');
            return sb.ToString();
        }

        private static string LoadEmbeddedHtml(string name)
        {
            var assembly = Assembly.GetExecutingAssembly();
            string html = LoadResource(assembly, "BioDraw.WebUI." + name);
            if (html == null)
            {
                return "<html><body><p style='font-family:sans-serif;padding:20px'>" +
                       "Resource not found.</p></body></html>";
            }
            string css = LoadResource(assembly, "BioDraw.WebUI.shared.css") ?? string.Empty;
            string js = LoadResource(assembly, "BioDraw.WebUI.bridge.js") ?? string.Empty;
            return html.Replace("STYLES_PLACEHOLDER", css)
                .Replace("BRIDGE_PLACEHOLDER", js);
        }

        private static string LoadResource(Assembly assembly, string name)
        {
            using (var stream = assembly.GetManifestResourceStream(name))
            {
                if (stream == null) return null;
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                    return reader.ReadToEnd();
            }
        }

        protected static string GetString(Dictionary<string, object> values, string key)
        {
            if (values == null) return string.Empty;
            object value;
            if (!values.TryGetValue(key, out value) || value == null) return string.Empty;
            return value as string ?? Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        protected static int GetInt(Dictionary<string, object> values, string key, int fallback)
        {
            if (values == null) return fallback;
            object value;
            if (!values.TryGetValue(key, out value) || value == null) return fallback;
            if (value is int) return (int)value;
            int parsed;
            return int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed)
                ? parsed
                : fallback;
        }

        protected static double GetDouble(Dictionary<string, object> values, string key, double fallback)
        {
            if (values == null) return fallback;
            object value;
            if (!values.TryGetValue(key, out value) || value == null) return fallback;
            if (value is double) return (double)value;
            if (value is decimal) return (double)(decimal)value;
            if (value is int) return (int)value;
            double parsed;
            return double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)
                ? parsed
                : fallback;
        }

        protected static bool GetBool(Dictionary<string, object> values, string key)
        {
            if (values == null) return false;
            object value;
            if (!values.TryGetValue(key, out value) || value == null) return false;
            if (value is bool) return (bool)value;
            bool parsed;
            return bool.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out parsed) && parsed;
        }

        protected static int ClampInt(int value, int minimum, int maximum)
        {
            return Math.Max(minimum, Math.Min(maximum, value));
        }

        protected static string NormalizeChoice(string value, string fallback, params string[] allowed)
        {
            foreach (var candidate in allowed)
            {
                if (string.Equals(value, candidate, StringComparison.OrdinalIgnoreCase))
                    return candidate;
            }
            return fallback;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && webView != null)
            {
                webView.NavigationCompleted -= OnNavigationCompleted;
                webView.WebMessageReceived -= OnWebMessageReceived;
                webView.Dispose();
                webView = null;
            }
            base.Dispose(disposing);
        }
    }
}
