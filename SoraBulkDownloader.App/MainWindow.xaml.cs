using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using Microsoft.Web.WebView2.Core;

namespace SoraBulkDownloader.App;

public partial class MainWindow : Window
{
    // ── State ────────────────────────────────────────────────────────────────
    private readonly ObservableCollection<DownloadItem> _queue = new();
    private string _outputBaseFolder = "";
    private bool _captureEnabled;
    private SectionCategory _scanSection = SectionCategory.Drafts;
    private bool _webViewReady;
    private CancellationTokenSource? _collectDraftUrlsCts;
    private CancellationTokenSource? _downloadCts;
    private bool _sessionExpired;

    // ── Tray ─────────────────────────────────────────────────────────────────
    private System.Windows.Forms.NotifyIcon? _trayIcon;

    // ── Persistence ──────────────────────────────────────────────────────────
    private static string QueuePersistPath => System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SoraBulkDownloader", "queue.json");

    // ── Scroll script (shared between collect loop and scan) ─────────────────
    private const string ScrollScript = @"
(() => {
  function findScrollables() {
    return Array.from(document.querySelectorAll('*')).filter(el => {
      const s = window.getComputedStyle(el);
      return s && (s.overflowY === 'auto' || s.overflowY === 'scroll') && el.scrollHeight > el.clientHeight + 20;
    }).slice(0, 30);
  }
  const els = findScrollables();
  try { window.scrollTo(0, document.body.scrollHeight); } catch {}
  for (const el of els) { try { el.scrollTop = el.scrollHeight; } catch {} }
})();";

    // ── Constructor ──────────────────────────────────────────────────────────
    public MainWindow()
    {
        InitializeComponent();
        QueueGrid.ItemsSource = _queue;
        _queue.CollectionChanged += (_, _) => UpdateQueueSummary();
        Loaded += MainWindow_OnLoaded;
        Closing += MainWindow_OnClosing;
        StateChanged += MainWindow_OnStateChanged;
    }

    // ════════════════════════════════════════════════════════════════════════
    // Startup / Shutdown
    // ════════════════════════════════════════════════════════════════════════

    private async void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
    {
        try
        {
            SetupTrayIcon();

            await Browser.EnsureCoreWebView2Async();
            _webViewReady = true;

            Browser.CoreWebView2.Settings.IsStatusBarEnabled = false;
            Browser.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            Browser.CoreWebView2.Settings.AreDevToolsEnabled = true;

            Browser.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);
            Browser.CoreWebView2.WebResourceResponseReceived += CoreWebView2_OnWebResourceResponseReceived;
            Browser.CoreWebView2.WebResourceRequested += CoreWebView2_OnWebResourceRequested;

            var downloads = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            _outputBaseFolder = System.IO.Path.Combine(downloads, "SoraBulkDownloader");
            OutputFolderText.Text = _outputBaseFolder;

            Browser.Source = new Uri("https://sora.chatgpt.com/drafts");

            await LoadPersistedQueueAsync();
            UpdateButtonStates();

            StatusText.Text = _queue.Count > 0
                ? $"Loaded {_queue.Count} pending item(s) from last session. Log in if needed."
                : "Log in if needed, then click Scan page.";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Failed to initialize: {ex.Message}";
        }
    }

    private void MainWindow_OnClosing(object? sender, CancelEventArgs e)
    {
        // Set Visible = false before disposing; otherwise the WinForms message pump
        // keeps the process alive after the WPF window is gone.
        if (_trayIcon != null)
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }
        SaveQueueSync();
        System.Windows.Application.Current.Shutdown();
    }

    private void MainWindow_OnStateChanged(object? sender, EventArgs e)
    {
        if (WindowState == WindowState.Minimized)
        {
            Hide();
            _trayIcon?.ShowBalloonTip(2000, "Sora Bulk Downloader",
                "Still running in the background. Double-click to restore.",
                System.Windows.Forms.ToolTipIcon.Info);
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // System tray
    // ════════════════════════════════════════════════════════════════════════

    private void SetupTrayIcon()
    {
        try
        {
            _trayIcon = new System.Windows.Forms.NotifyIcon
            {
                Icon = System.Drawing.SystemIcons.Application,
                Text = "Sora Bulk Downloader",
                Visible = true,
            };
            _trayIcon.DoubleClick += (_, _) => RestoreFromTray();

            var menu = new System.Windows.Forms.ContextMenuStrip();
            menu.Items.Add("Restore", null, (_, _) => RestoreFromTray());
            menu.Items.Add("Open Output Folder", null, (_, _) => OpenOutputFolder());
            menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
            menu.Items.Add("Exit", null, (_, _) =>
            {
                _trayIcon.Visible = false;
                System.Windows.Application.Current.Shutdown();
            });
            _trayIcon.ContextMenuStrip = menu;
        }
        catch { /* tray icon is optional */ }
    }

    private void RestoreFromTray()
    {
        // Must run on UI thread (called from WinForms tray event)
        Dispatcher.InvokeAsync(() =>
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
        });
    }

    // ════════════════════════════════════════════════════════════════════════
    // Network capture
    // ════════════════════════════════════════════════════════════════════════

    private void CoreWebView2_OnWebResourceResponseReceived(
        object? sender, CoreWebView2WebResourceResponseReceivedEventArgs e)
    {
        try
        {
            if (!_captureEnabled) return;
            var uriStr = e.Request.Uri.ToLowerInvariant();
            var ct = (e.Response.Headers.GetHeader("content-type") ?? "").ToLowerInvariant();

            var isImage = ct.StartsWith("image/") || ct.Contains("image") ||
                          uriStr.EndsWith(".png") || uriStr.EndsWith(".jpg") ||
                          uriStr.EndsWith(".jpeg") || uriStr.EndsWith(".webp") ||
                          uriStr.EndsWith(".gif") || uriStr.Contains("/thumb");

            var isVideo = ct.Contains("video") || ct.Contains("mpegurl") ||
                          uriStr.EndsWith(".mp4") || uriStr.EndsWith(".webm") ||
                          uriStr.EndsWith(".m3u8") || uriStr.EndsWith(".ts") ||
                          uriStr.EndsWith(".m4s") || uriStr.EndsWith(".mpd") ||
                          uriStr.Contains(".mp4?") || uriStr.Contains(".webm?");

            if (!isVideo && !isImage) return;
            TryEnqueueUrl(e.Request.Uri, isVideo ? DownloadKind.Video : DownloadKind.Thumbnail);
        }
        catch { }
    }

    private void CoreWebView2_OnWebResourceRequested(
        object? sender, CoreWebView2WebResourceRequestedEventArgs e)
    {
        try
        {
            if (!_captureEnabled) return;
            if (!TryClassifyMediaUrl(e.Request.Uri, out var kind)) return;
            TryEnqueueUrl(e.Request.Uri, kind);
        }
        catch { }
    }

    private void TryEnqueueUrl(string url, DownloadKind kind)
    {
        if (kind == DownloadKind.Video && IncludeVideosCheck.IsChecked != true) return;
        if (kind == DownloadKind.Thumbnail && IncludeThumbsCheck.IsChecked != true) return;

        if (_queue.Any(q => q.Kind == kind && q.Section == _scanSection &&
                            StringComparer.OrdinalIgnoreCase.Equals(q.Url, url)))
            return;

        _queue.Add(new DownloadItem
        {
            Kind = kind,
            Section = _scanSection,
            Url = url,
            Status = DownloadStatus.Found
        });
        StatusText.Text = $"Captured: {_queue.Count} items queued";
        UpdateButtonStates();
    }

    private static bool TryClassifyMediaUrl(string url, out DownloadKind kind)
    {
        kind = DownloadKind.Video;
        var u = url.ToLowerInvariant();
        if (u.Contains("/thumb") || u.Contains("thumb") || u.EndsWith(".png") || u.EndsWith(".jpg") ||
            u.EndsWith(".jpeg") || u.EndsWith(".webp") || u.EndsWith(".gif") || u.EndsWith(".svg"))
        {
            kind = DownloadKind.Thumbnail;
            return true;
        }
        if (u.Contains("video") || u.Contains("/video") || u.Contains("mp4") || u.Contains("webm") ||
            u.Contains("m3u8") || u.Contains("mpegurl") || u.Contains(".ts") || u.Contains(".m4s") ||
            u.Contains(".mpd") || u.Contains("hls") || u.Contains("/media"))
        {
            kind = DownloadKind.Video;
            return true;
        }
        return false;
    }

    // ════════════════════════════════════════════════════════════════════════
    // Navigation toolbar buttons
    // ════════════════════════════════════════════════════════════════════════

    private void GoDraftsButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!_webViewReady) return;
        Browser.Source = new Uri("https://sora.chatgpt.com/drafts");
    }

    private void GoProfileButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!_webViewReady) return;
        Browser.Source = new Uri("https://sora.chatgpt.com/profile");
    }

    private void PickFolderButton_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Choose where to save your Sora downloads",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
            SelectedPath = _outputBaseFolder
        };
        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
        _outputBaseFolder = dialog.SelectedPath;
        OutputFolderText.Text = _outputBaseFolder;
    }

    private void OpenFolderButton_OnClick(object sender, RoutedEventArgs e) => OpenOutputFolder();

    private void OpenOutputFolder()
    {
        if (string.IsNullOrWhiteSpace(_outputBaseFolder)) return;
        try
        {
            System.IO.Directory.CreateDirectory(_outputBaseFolder);
            Process.Start("explorer.exe", _outputBaseFolder);
        }
        catch { }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Collect Draft URLs (loop)
    // ════════════════════════════════════════════════════════════════════════

    private async void CollectDraftUrlsButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!_webViewReady) return;

        // Toggle: click again to stop
        if (_collectDraftUrlsCts is not null)
        {
            _collectDraftUrlsCts.Cancel();
            _collectDraftUrlsCts = null;
            StatusText.Text = $"Stopped. Total queued: {_queue.Count}.";
            return;
        }

        _collectDraftUrlsCts = new CancellationTokenSource();
        var token = _collectDraftUrlsCts.Token;
        CollectDraftUrlsButton.Content = "⏹ Stop collecting";
        SetNavigationEnabled(false);

        try
        {
            _scanSection = SectionCategory.Drafts;
            Browser.Source = new Uri("https://sora.chatgpt.com/drafts");
            await WaitForScanPageReadyAsync(SectionCategory.Drafts);
            _captureEnabled = true;
            var noNewRounds = 0;

            for (var round = 1; round <= 200; round++)
            {
                token.ThrowIfCancellationRequested();
                var before = _queue.Count;
                StatusText.Text = $"Collecting… round {round} — {before} queued so far";
                await Browser.ExecuteScriptAsync(ScrollScript);
                await Task.Delay(1500, token);
                var added = _queue.Count - before;
                noNewRounds = added <= 0 ? noNewRounds + 1 : 0;
                if (noNewRounds >= 10) break;
            }

            StatusText.Text = $"Collection done. {_queue.Count} items queued. Click Export or Download.";
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = $"Stopped. {_queue.Count} items queued.";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Collection failed: {ex.Message}";
        }
        finally
        {
            _captureEnabled = false;
            _collectDraftUrlsCts = null;
            CollectDraftUrlsButton.Content = "Collect Draft URLs (loop)";
            SetNavigationEnabled(true);
            UpdateButtonStates();
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Scan
    // ════════════════════════════════════════════════════════════════════════

    private async void ScanButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!_webViewReady) return;
        ScanButton.IsEnabled = false;
        try
        {
            var section = InferSectionFromCurrentUrl();
            StatusText.Text = $"Scanning page… ({GetSectionLabel(section)})";
            await WaitForScanPageReadyAsync(section);
            await ScanAndQueueAsync(section);
        }
        finally
        {
            ScanButton.IsEnabled = true;
            UpdateButtonStates();
        }
    }

    /// <summary>Clears the queue, scrolls the page, then enqueues all found media URLs.</summary>
    private async Task<int> ScanAndQueueAsync(SectionCategory section)
    {
        _scanSection = section;
        _queue.Clear();
        _captureEnabled = true;
        StatusText.Text = $"Scanning {GetSectionLabel(section)}…";
        try
        {
            return await RunScanScriptAsync();
        }
        finally
        {
            _captureEnabled = false;
            StatusText.Text = $"Scan done — {_queue.Count} items queued for {GetSectionLabel(section)}.";
            UpdateButtonStates();
        }
    }

    private async Task<int> RunScanScriptAsync()
    {
        const string scanScript = @"
(async () => {
  function uniq(a) { return Array.from(new Set(a.filter(Boolean))); }
  function findScrollables() {
    return Array.from(document.querySelectorAll('*')).filter(el => {
      const s = window.getComputedStyle(el);
      return s && (s.overflowY === 'auto' || s.overflowY === 'scroll') && el.scrollHeight > el.clientHeight + 20;
    }).slice(0, 30);
  }
  let scrollables = findScrollables();
  function scrollAll() {
    try { window.scrollTo(0, document.body.scrollHeight); } catch {}
    for (const el of scrollables) { try { el.scrollTop = el.scrollHeight; } catch {} }
  }
  let lastH = 0, stable = 0;
  for (let i = 0; i < 250; i++) {
    if (i % 15 === 0) scrollables = findScrollables();
    scrollAll();
    await new Promise(r => setTimeout(r, 700));
    const heights = [document.body ? document.body.scrollHeight : 0,
                     ...scrollables.map(el => el.scrollHeight || 0)];
    const h = Math.max(...heights, 0);
    stable = (h === lastH) ? stable + 1 : 0;
    lastH = h;
    let atBottom = (window.innerHeight + window.scrollY) >= (h - 5);
    for (const el of scrollables)
      if (el.scrollTop + el.clientHeight >= el.scrollHeight - 5) atBottom = true;
    if (atBottom && stable >= 8) break;
  }
  const vids = [];
  for (const v of document.querySelectorAll('video')) {
    if (v.currentSrc) vids.push(v.currentSrc);
    if (v.src) vids.push(v.src);
  }
  for (const s of document.querySelectorAll('video source')) if (s.src) vids.push(s.src);
  const imgs = [];
  for (const img of document.querySelectorAll('img')) {
    if (img.currentSrc) imgs.push(img.currentSrc);
    if (img.src) imgs.push(img.src);
  }
  const links = [];
  for (const a of document.querySelectorAll('a[href]')) {
    const h = a.getAttribute('href');
    if (!h) continue;
    links.push(h.startsWith('/') ? location.origin + h : h.startsWith('http') ? h : '');
  }
  return JSON.stringify({ videoUrls: uniq(vids), thumbUrls: uniq(imgs), pageLinks: uniq(links) });
})();";

        try
        {
            var raw = await Browser.ExecuteScriptAsync(scanScript);
            var result = JsonSerializer.Deserialize<ScanResult>(UnwrapWebView2JsonString(raw));
            if (result is null) return 0;

            var added = 0;
            if (IncludeVideosCheck.IsChecked == true)
            {
                foreach (var u in result.videoUrls.Where(IsHttpUrl))
                {
                    if (_queue.Any(q => q.Kind == DownloadKind.Video && q.Section == _scanSection &&
                                       StringComparer.OrdinalIgnoreCase.Equals(q.Url, u)))
                        continue;
                    _queue.Add(new DownloadItem { Kind = DownloadKind.Video, Section = _scanSection, Url = u, Status = DownloadStatus.Found });
                    added++;
                }
            }
            if (IncludeThumbsCheck.IsChecked == true)
            {
                foreach (var u in result.thumbUrls.Where(IsHttpUrl))
                {
                    if (_queue.Any(q => q.Kind == DownloadKind.Thumbnail && q.Section == _scanSection &&
                                       StringComparer.OrdinalIgnoreCase.Equals(q.Url, u)))
                        continue;
                    _queue.Add(new DownloadItem { Kind = DownloadKind.Thumbnail, Section = _scanSection, Url = u, Status = DownloadStatus.Found });
                    added++;
                }
            }
            return added;
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Scan error: {ex.Message}";
            return 0;
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Download All
    // ════════════════════════════════════════════════════════════════════════

    private async void DownloadAllButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!_webViewReady || string.IsNullOrWhiteSpace(_outputBaseFolder)) return;

        SetNavigationEnabled(false);
        DownloadAllButton.IsEnabled = false;
        _downloadCts = new CancellationTokenSource();
        CancelDownloadButton.IsEnabled = true;
        _sessionExpired = false;
        ResumeDownloadButton.Visibility = Visibility.Collapsed;
        OverallProgress.Value = 0;

        var totalDone = 0;
        var totalFailed = 0;

        try
        {
            var sections = new[]
            {
                (SectionCategory.Drafts, "Drafts"),
                (SectionCategory.Likes,  "Likes"),
                (SectionCategory.Posts,  "Posts"),
            };

            foreach (var (section, label) in sections)
            {
                if (_downloadCts.IsCancellationRequested) break;

                StatusText.Text = $"Navigating to {label}…";
                await NavigateToSectionAsync(section);
                if (_downloadCts.IsCancellationRequested) break;

                StatusText.Text = $"Scanning {label}…";
                var added = await ScanAndQueueAsync(section);
                if (added == 0)
                {
                    StatusText.Text = $"No media found in {label}. (Are you logged in?)";
                    continue;
                }

                StatusText.Text = $"Downloading {label}… ({added} items)";
                await DownloadQueueAsync(_downloadCts.Token);

                totalDone += _queue.Count(q => q.Status == DownloadStatus.Done);
                totalFailed += _queue.Count(q => q.Status == DownloadStatus.Failed);

                if (_sessionExpired) break;
            }

            if (!_sessionExpired)
            {
                var summary = $"All done! {totalDone} downloaded, {totalFailed} failed.";
                StatusText.Text = summary;
                OverallProgress.Value = 100;

                try { System.Media.SystemSounds.Asterisk.Play(); } catch { }
                _trayIcon?.ShowBalloonTip(6000, "Sora Bulk Downloader", summary,
                    System.Windows.Forms.ToolTipIcon.Info);

                if (totalFailed == 0)
                {
                    if (System.Windows.MessageBox.Show($"{summary}\n\nOpen output folder?",
                            "Download Complete", MessageBoxButton.YesNo, MessageBoxImage.Information)
                        == MessageBoxResult.Yes)
                    {
                        OpenOutputFolder();
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show(
                        $"{summary}\n\nClick 'Retry Failed' to retry the {totalFailed} failed item(s).",
                        "Download Complete", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = $"Cancelled. {totalDone} downloaded so far.";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Download all failed: {ex.Message}";
        }
        finally
        {
            _downloadCts?.Dispose();
            _downloadCts = null;
            SetNavigationEnabled(true);
            DownloadAllButton.IsEnabled = true;
            CancelDownloadButton.IsEnabled = false;
            UpdateButtonStates();
            await SaveQueueAsync();
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Download / Cancel / Retry / Resume buttons
    // ════════════════════════════════════════════════════════════════════════

    private async void DownloadButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!_webViewReady || string.IsNullOrWhiteSpace(_outputBaseFolder))
        {
            if (string.IsNullOrWhiteSpace(_outputBaseFolder))
                StatusText.Text = "Pick an output folder first.";
            return;
        }
        if (!_queue.Any(q => q.Status is DownloadStatus.Found or DownloadStatus.Failed))
        {
            StatusText.Text = "Nothing to download — scan a page first.";
            return;
        }

        DownloadButton.IsEnabled = false;
        ScanButton.IsEnabled = false;
        _downloadCts = new CancellationTokenSource();
        CancelDownloadButton.IsEnabled = true;
        _sessionExpired = false;
        ResumeDownloadButton.Visibility = Visibility.Collapsed;

        try
        {
            await DownloadQueueAsync(_downloadCts.Token);
        }
        finally
        {
            _downloadCts?.Dispose();
            _downloadCts = null;
            DownloadButton.IsEnabled = true;
            ScanButton.IsEnabled = true;
            CancelDownloadButton.IsEnabled = false;
            UpdateButtonStates();
            await SaveQueueAsync();
        }
    }

    private void CancelDownloadButton_OnClick(object sender, RoutedEventArgs e)
    {
        _downloadCts?.Cancel();
        StatusText.Text = "Cancelling…";
        CancelDownloadButton.IsEnabled = false;
    }

    private void RetryFailedButton_OnClick(object sender, RoutedEventArgs e)
    {
        var failed = _queue.Where(q => q.Status == DownloadStatus.Failed).ToList();
        if (failed.Count == 0) return;

        foreach (var item in failed)
        {
            item.Status = DownloadStatus.Found;
            item.ProgressText = "";
            item.Speed = "";
        }

        StatusText.Text = $"Re-queued {failed.Count} failed items.";
        UpdateButtonStates();

        // Kick off immediately
        DownloadButton_OnClick(sender, e);
    }

    private void ResumeDownloadButton_OnClick(object sender, RoutedEventArgs e)
    {
        _sessionExpired = false;
        ResumeDownloadButton.Visibility = Visibility.Collapsed;
        StatusText.Text = "Resuming downloads…";
        DownloadButton_OnClick(sender, e);
    }

    // ════════════════════════════════════════════════════════════════════════
    // Core download engine — concurrent with retry, speed tracking, auth detection
    // ════════════════════════════════════════════════════════════════════════

    private async Task DownloadQueueAsync(CancellationToken ct)
    {
        var pending = _queue.Where(q => q.Status is DownloadStatus.Found or DownloadStatus.Failed).ToList();
        if (pending.Count == 0) return;

        var cookieContainer = await ExportCookiesAsync();
        using var handler = new HttpClientHandler
        {
            CookieContainer = cookieContainer,
            AutomaticDecompression =
                DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.Brotli
        };
        using var http = new HttpClient(handler) { Timeout = TimeSpan.FromMinutes(15) };
        http.DefaultRequestHeaders.UserAgent.ParseAdd(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) SoraBulkDownloader/2.0");

        // Up to 4 concurrent downloads
        var semaphore = new SemaphoreSlim(4);
        var doneCount = 0;
        var startTime = DateTime.UtcNow;

        var tasks = pending.Select(async item =>
        {
            if (ct.IsCancellationRequested)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    if (item.Status == DownloadStatus.Downloading)
                        item.Status = DownloadStatus.Found;
                });
                return;
            }

            await semaphore.WaitAsync(ct);
            try
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    item.Status = DownloadStatus.Downloading;
                    item.ProgressText = "starting…";
                    item.Speed = "";
                });

                var sectionFolder = GetSectionOutputFolder(item.Section, item.Kind);
                await DownloadWithRetryAsync(http, item, sectionFolder, ct);

                var dc = Interlocked.Increment(ref doneCount);
                var elapsed = (DateTime.UtcNow - startTime).TotalSeconds;
                await Dispatcher.InvokeAsync(() =>
                {
                    OverallProgress.Value = (double)dc / pending.Count * 100.0;
                    var eta = dc < pending.Count && elapsed > 0
                        ? $" — ETA {FormatEta(TimeSpan.FromSeconds((pending.Count - dc) / (dc / elapsed)))}"
                        : "";
                    StatusText.Text = $"Downloaded {dc}/{pending.Count}{eta}";
                });
            }
            catch (OperationCanceledException)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    if (item.Status == DownloadStatus.Downloading)
                        item.Status = DownloadStatus.Found;
                });
            }
            catch (SessionExpiredException)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    item.Status = DownloadStatus.Failed;
                    item.ProgressText = "Session expired";
                    item.Speed = "";
                    _sessionExpired = true;
                    ResumeDownloadButton.Visibility = Visibility.Visible;
                    StatusText.Text =
                        "⚠️ Session expired — log in again in the browser above, then click Resume.";
                });
                _downloadCts?.Cancel();
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    item.Status = DownloadStatus.Failed;
                    item.ProgressText = ex.Message.Length > 80
                        ? ex.Message[..77] + "…"
                        : ex.Message;
                    item.Speed = "";
                });
            }
            finally
            {
                semaphore.Release();
            }
        });

        await Task.WhenAll(tasks);

        var failedCount = _queue.Count(q => q.Status == DownloadStatus.Failed);
        await Dispatcher.InvokeAsync(() =>
        {
            UpdateButtonStates();
            if (failedCount > 0 && !_sessionExpired)
                StatusText.Text += $"  ({failedCount} failed — click 'Retry Failed')";
        });
    }

    private async Task DownloadWithRetryAsync(
        HttpClient http, DownloadItem item, string sectionFolder, CancellationToken ct)
    {
        const int maxAttempts = 3;
        Exception? lastEx = null;

        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                await DownloadToFolderAsync(http, item.Url, sectionFolder, item.Kind,
                    async (progressText, speed) =>
                    {
                        await Dispatcher.InvokeAsync(() =>
                        {
                            item.ProgressText = progressText;
                            item.Speed = speed;
                        });
                    }, ct);

                await Dispatcher.InvokeAsync(() =>
                {
                    item.Status = DownloadStatus.Done;
                    item.ProgressText = "100%";
                    item.Speed = "";
                });
                return;
            }
            catch (HttpRequestException ex)
                when (ex.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
            {
                throw new SessionExpiredException();
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                lastEx = ex;
                if (attempt < maxAttempts - 1)
                {
                    var delay = TimeSpan.FromSeconds(Math.Pow(2, attempt)); // 1s, 2s
                    await Dispatcher.InvokeAsync(() =>
                        item.ProgressText = $"Retry {attempt + 1}/{maxAttempts - 1} in {delay.TotalSeconds:0}s…");
                    await Task.Delay(delay, ct);
                }
            }
        }

        throw lastEx!;
    }

    private static async Task DownloadToFolderAsync(
        HttpClient http,
        string url,
        string outputFolder,
        DownloadKind kind,
        Func<string, string, Task>? onProgress = null,
        CancellationToken ct = default)
    {
        using var resp = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
        resp.EnsureSuccessStatusCode();

        var contentType = resp.Content.Headers.ContentType?.MediaType ?? "";
        var ext = GuessExtension(kind, contentType, url);
        var safeKind = kind == DownloadKind.Video ? "video" : "thumb";

        System.IO.Directory.CreateDirectory(outputFolder);
        var hash = Sha256Hex(url);
        var fileName = $"{safeKind}_{hash}{ext}";
        var fullPath = System.IO.Path.Combine(outputFolder, fileName);

        // Skip if already fully downloaded
        if (System.IO.File.Exists(fullPath) && new System.IO.FileInfo(fullPath).Length > 0)
        {
            if (onProgress != null)
                await onProgress("already saved", "");
            return;
        }

        var totalBytes = resp.Content.Headers.ContentLength ?? -1L;
        long bytesRead = 0;
        var sw = Stopwatch.StartNew();
        long lastSpeedMs = 0;
        long lastSpeedBytes = 0;

        await using var stream = await resp.Content.ReadAsStreamAsync(ct);
        // Write to a temp file first; rename on completion to avoid partial files
        var tmpPath = fullPath + ".tmp";
        try
        {
            await using (var fs = System.IO.File.Create(tmpPath))
            {
                var buffer = new byte[81920]; // 80 KB
                int read;
                while ((read = await stream.ReadAsync(buffer, ct)) > 0)
                {
                    await fs.WriteAsync(buffer.AsMemory(0, read), ct);
                    bytesRead += read;

                    var nowMs = sw.ElapsedMilliseconds;
                    if (onProgress != null && nowMs - lastSpeedMs >= 500)
                    {
                        var deltaMs = nowMs - lastSpeedMs;
                        var speed = deltaMs > 0
                            ? (bytesRead - lastSpeedBytes) * 1000.0 / deltaMs
                            : 0.0;
                        lastSpeedMs = nowMs;
                        lastSpeedBytes = bytesRead;

                        var progressText = totalBytes > 0
                            ? $"{bytesRead * 100 / totalBytes}%  {FormatBytes(bytesRead)}/{FormatBytes(totalBytes)}"
                            : FormatBytes(bytesRead);

                        await onProgress(progressText, FormatSpeed(speed));
                    }
                }
            }

            // Atomic rename
            if (System.IO.File.Exists(fullPath)) System.IO.File.Delete(fullPath);
            System.IO.File.Move(tmpPath, fullPath);
        }
        catch
        {
            // Clean up temp file on failure
            try { if (System.IO.File.Exists(tmpPath)) System.IO.File.Delete(tmpPath); } catch { }
            throw;
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Export URLs
    // ════════════════════════════════════════════════════════════════════════

    private void ExportUrlsButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_outputBaseFolder))
            {
                StatusText.Text = "Pick an output folder first.";
                return;
            }
            if (_queue.Count == 0)
            {
                StatusText.Text = "Queue is empty — scan a page first.";
                return;
            }

            var exportDir = System.IO.Path.Combine(_outputBaseFolder, "exports");
            System.IO.Directory.CreateDirectory(exportDir);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var items = _queue.ToList();
            var payload = items.Select(x => new
            {
                section = GetSectionLabel(x.Section),
                kind = x.Kind.ToString(),
                status = x.Status.ToString(),
                url = x.Url,
                file = x.FilePath,
                exportedAt = DateTime.Now.ToString("o")
            }).ToList();

            // JSON
            var jsonPath = System.IO.Path.Combine(exportDir, $"urls_{stamp}.json");
            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
            System.IO.File.WriteAllText(jsonPath, json, Encoding.UTF8);

            // CSV
            var csvPath = System.IO.Path.Combine(exportDir, $"urls_{stamp}.csv");
            var csv = new StringBuilder("section,kind,status,url,file,exportedAt\r\n");
            foreach (var p in payload)
                csv.AppendLine(
                    $"{CsvEscape(p.section)},{CsvEscape(p.kind)},{CsvEscape(p.status)}," +
                    $"{CsvEscape(p.url)},{CsvEscape(p.file ?? "")},{CsvEscape(p.exportedAt)}");
            System.IO.File.WriteAllText(csvPath, csv.ToString(), Encoding.UTF8);

            // Clipboard — plain URLs
            var plainUrls = string.Join(Environment.NewLine, payload.Select(p => p.url));
            if (!string.IsNullOrWhiteSpace(plainUrls))
                System.Windows.Clipboard.SetText(plainUrls);

            StatusText.Text = $"Exported {items.Count} items to: {exportDir}  (URLs copied to clipboard)";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Export failed: {ex.Message}";
        }
    }

    private static string CsvEscape(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }

    // ════════════════════════════════════════════════════════════════════════
    // Queue persistence
    // ════════════════════════════════════════════════════════════════════════

    private async Task SaveQueueAsync()
    {
        try
        {
            // Only persist items that still need work
            var toSave = _queue
                .Where(q => q.Status != DownloadStatus.Done)
                .Select(q => new PersistedItem(
                    GetSectionLabel(q.Section),
                    q.Kind.ToString(),
                    q.Url,
                    q.Status.ToString(),
                    q.FilePath))
                .ToList();

            var dir = System.IO.Path.GetDirectoryName(QueuePersistPath)!;
            System.IO.Directory.CreateDirectory(dir);
            var json = JsonSerializer.Serialize(toSave, new JsonSerializerOptions { WriteIndented = true });
            await System.IO.File.WriteAllTextAsync(QueuePersistPath, json, Encoding.UTF8);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Warning: could not save queue — {ex.Message}";
        }
    }

    private void SaveQueueSync()
    {
        try
        {
            var toSave = _queue
                .Where(q => q.Status != DownloadStatus.Done)
                .Select(q => new PersistedItem(
                    GetSectionLabel(q.Section),
                    q.Kind.ToString(),
                    q.Url,
                    q.Status.ToString(),
                    q.FilePath))
                .ToList();

            var dir = System.IO.Path.GetDirectoryName(QueuePersistPath)!;
            System.IO.Directory.CreateDirectory(dir);
            var json = JsonSerializer.Serialize(toSave, new JsonSerializerOptions { WriteIndented = true });
            System.IO.File.WriteAllText(QueuePersistPath, json, Encoding.UTF8);
        }
        catch { }
    }

    private async Task LoadPersistedQueueAsync()
    {
        try
        {
            if (!System.IO.File.Exists(QueuePersistPath)) return;
            var json = await System.IO.File.ReadAllTextAsync(QueuePersistPath, Encoding.UTF8);
            var items = JsonSerializer.Deserialize<List<PersistedItem>>(json);
            if (items is null || items.Count == 0) return;

            foreach (var item in items.Where(i => i.Status != "Done"))
            {
                if (_queue.Any(q => StringComparer.OrdinalIgnoreCase.Equals(q.Url, item.Url)))
                    continue;

                var section = item.Section switch
                {
                    "Likes" => SectionCategory.Likes,
                    "Posts" => SectionCategory.Posts,
                    _ => SectionCategory.Drafts
                };
                var kind = item.Kind == "Thumbnail" ? DownloadKind.Thumbnail : DownloadKind.Video;
                var status = item.Status == "Failed" ? DownloadStatus.Failed : DownloadStatus.Found;

                _queue.Add(new DownloadItem
                {
                    Section = section,
                    Kind = kind,
                    Url = item.Url,
                    Status = status,
                    FilePath = item.FilePath
                });
            }
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Warning: could not load saved queue — {ex.Message}";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Page navigation helpers
    // ════════════════════════════════════════════════════════════════════════

    private async Task NavigateToSectionAsync(SectionCategory section)
    {
        if (section == SectionCategory.Drafts)
        {
            Browser.Source = new Uri("https://sora.chatgpt.com/drafts");
            await WaitForScanPageReadyAsync(SectionCategory.Drafts);
            return;
        }

        // Try the direct sub-URL first (Sora may support /profile/likes and /profile/posts)
        var directUrl = section == SectionCategory.Likes
            ? "https://sora.chatgpt.com/profile/likes"
            : "https://sora.chatgpt.com/profile/posts";

        Browser.Source = new Uri(directUrl);
        await WaitForScanPageReadyAsync(section);

        // If we ended up on a generic profile page without the right tab, click it
        var landed = Browser.Source?.AbsoluteUri?.ToLowerInvariant() ?? "";
        var keyword = section == SectionCategory.Likes ? "likes" : "posts";
        if (!landed.Contains(keyword))
        {
            Browser.Source = new Uri("https://sora.chatgpt.com/profile");
            await WaitForScanPageReadyAsync(section);
            await ClickProfileTabAsync(section == SectionCategory.Likes ? "Likes" : "Posts");
            await WaitForScanPageReadyAsync(section);
        }
    }

    private async Task WaitForScanPageReadyAsync(SectionCategory section)
    {
        var sw = Stopwatch.StartNew();
        long? lastScrollHeight = null;
        int stableTicks = 0;

        while (sw.ElapsedMilliseconds < 60_000)
        {
            try
            {
                var raw = await Browser.ExecuteScriptAsync(
                    "JSON.stringify({url:location.href,videos:document.querySelectorAll('video').length," +
                    "scrollHeight:document.body?document.body.scrollHeight:0});");
                var probe = JsonSerializer.Deserialize<ProbeResult>(UnwrapWebView2JsonString(raw));
                if (probe is null) break;

                var url = (probe.url ?? "").ToLowerInvariant();
                var urlOk = section == SectionCategory.Drafts
                    ? url.Contains("/drafts")
                    : url.Contains("/profile");

                if (urlOk && (probe.videos > 0 ||
                              (lastScrollHeight.HasValue && probe.scrollHeight > lastScrollHeight.Value + 500)))
                    return;

                stableTicks = lastScrollHeight.HasValue && probe.scrollHeight == lastScrollHeight.Value
                    ? stableTicks + 1
                    : 0;
                lastScrollHeight = probe.scrollHeight;

                if (urlOk && stableTicks >= 8) return;
            }
            catch { }

            await Task.Delay(800);
        }
    }

    private async Task ClickProfileTabAsync(string tabLabel)
    {
        var labelLit = JsonSerializer.Serialize(tabLabel);
        var js = $@"(() => {{
  const label = {labelLit};
  const norm = s => (s||'').trim().toLowerCase();
  const t = norm(label);
  const all = Array.from(document.querySelectorAll('a,button,[role=""tab""],div,span'));
  const exact = all.find(el => norm(el.innerText) === t);
  if (exact) {{ exact.click(); return true; }}
  const inc = all.find(el => norm(el.innerText).includes(t));
  if (inc) {{ inc.click(); return true; }}
  const href = Array.from(document.querySelectorAll('a[href]'))
    .find(a => (a.getAttribute('href')||'').toLowerCase().includes(t));
  if (href) {{ href.click(); return true; }}
  return false;
}})();";
        try { await Browser.ExecuteScriptAsync(js); } catch { }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Cookies
    // ════════════════════════════════════════════════════════════════════════

    private async Task<CookieContainer> ExportCookiesAsync()
    {
        var container = new CookieContainer();
        try
        {
            var cookies = await Browser.CoreWebView2.CookieManager.GetCookiesAsync("https://sora.chatgpt.com");
            foreach (var c in cookies)
            {
                var cookie = new Cookie(c.Name, c.Value, c.Path,
                    c.Domain.StartsWith('.') ? c.Domain[1..] : c.Domain)
                {
                    Secure = c.IsSecure,
                    HttpOnly = c.IsHttpOnly,
                    Expires = c.Expires
                };
                container.Add(cookie);
            }
        }
        catch { }
        return container;
    }

    // ════════════════════════════════════════════════════════════════════════
    // UI helpers
    // ════════════════════════════════════════════════════════════════════════

    private void UpdateButtonStates()
    {
        var hasPending = _queue.Any(q => q.Status is DownloadStatus.Found or DownloadStatus.Failed);
        var hasFailed = _queue.Any(q => q.Status == DownloadStatus.Failed);
        DownloadButton.IsEnabled = hasPending;
        RetryFailedButton.IsEnabled = hasFailed;
        UpdateQueueSummary();
    }

    private void UpdateQueueSummary()
    {
        var total = _queue.Count;
        var done = _queue.Count(q => q.Status == DownloadStatus.Done);
        var failed = _queue.Count(q => q.Status == DownloadStatus.Failed);
        var pending = _queue.Count(q => q.Status == DownloadStatus.Found);
        var downloading = _queue.Count(q => q.Status == DownloadStatus.Downloading);

        var parts = new List<string>();
        if (total > 0) parts.Add($"{total} total");
        if (done > 0) parts.Add($"{done} done");
        if (downloading > 0) parts.Add($"{downloading} downloading");
        if (pending > 0) parts.Add($"{pending} pending");
        if (failed > 0) parts.Add($"{failed} failed");

        QueueSummaryText.Text = parts.Count > 0 ? $"({string.Join(" · ", parts)})" : "";
    }

    private void SetNavigationEnabled(bool enabled)
    {
        GoDraftsButton.IsEnabled = enabled;
        GoProfileButton.IsEnabled = enabled;
        ScanButton.IsEnabled = enabled;
        DownloadButton.IsEnabled = enabled;
        DownloadAllButton.IsEnabled = enabled;
    }

    // ════════════════════════════════════════════════════════════════════════
    // Misc helpers
    // ════════════════════════════════════════════════════════════════════════

    private static string GetSectionLabel(SectionCategory section) => section switch
    {
        SectionCategory.Drafts => "Drafts",
        SectionCategory.Likes => "Likes",
        SectionCategory.Posts => "Posts",
        _ => "Unknown"
    };

    private SectionCategory InferSectionFromCurrentUrl()
    {
        try
        {
            var uri = Browser.Source?.AbsoluteUri?.ToLowerInvariant() ?? "";
            if (uri.Contains("/drafts")) return SectionCategory.Drafts;
            if (uri.Contains("posts")) return SectionCategory.Posts;
            if (uri.Contains("likes") || uri.Contains("/profile")) return SectionCategory.Likes;
        }
        catch { }
        return SectionCategory.Drafts;
    }

    private string GetSectionOutputFolder(SectionCategory section, DownloadKind kind) =>
        System.IO.Path.Combine(
            _outputBaseFolder,
            GetSectionLabel(section),
            kind == DownloadKind.Video ? "video" : "thumb");

    private static string GuessExtension(DownloadKind kind, string contentType, string url)
    {
        var pathExt = System.IO.Path.GetExtension(new Uri(url).AbsolutePath);
        if (!string.IsNullOrWhiteSpace(pathExt) && pathExt.Length is >= 2 and <= 6)
            return pathExt;

        var u = url.ToLowerInvariant();
        if (u.Contains(".mp4"))  return ".mp4";
        if (u.Contains(".webm")) return ".webm";
        if (u.Contains(".m3u8")) return ".m3u8";
        if (u.Contains(".m4s"))  return ".m4s";

        if (kind == DownloadKind.Video)
        {
            if (contentType.Contains("mp4",      StringComparison.OrdinalIgnoreCase)) return ".mp4";
            if (contentType.Contains("webm",     StringComparison.OrdinalIgnoreCase)) return ".webm";
            if (contentType.Contains("mpegurl",  StringComparison.OrdinalIgnoreCase)) return ".m3u8";
            return ".mp4";
        }

        if (contentType.Contains("png",  StringComparison.OrdinalIgnoreCase)) return ".png";
        if (contentType.Contains("webp", StringComparison.OrdinalIgnoreCase)) return ".webp";
        if (contentType.Contains("jpeg", StringComparison.OrdinalIgnoreCase) ||
            contentType.Contains("jpg",  StringComparison.OrdinalIgnoreCase)) return ".jpg";
        return ".jpg";
    }

    private static string Sha256Hex(string input)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    private static bool IsHttpUrl(string url) =>
        url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
        url.StartsWith("http://",  StringComparison.OrdinalIgnoreCase);

    private static string UnwrapWebView2JsonString(string raw)
    {
        try { return JsonSerializer.Deserialize<string>(raw) ?? raw; }
        catch { return raw; }
    }

    // ── Formatting ────────────────────────────────────────────────────────────
    private static string FormatBytes(long bytes) => bytes switch
    {
        >= 1_073_741_824 => $"{bytes / 1_073_741_824.0:0.0} GB",
        >= 1_048_576     => $"{bytes / 1_048_576.0:0.0} MB",
        >= 1_024         => $"{bytes / 1_024.0:0.0} KB",
        _                => $"{bytes} B"
    };

    private static string FormatSpeed(double bytesPerSec) => bytesPerSec switch
    {
        >= 1_048_576 => $"{bytesPerSec / 1_048_576.0:0.0} MB/s",
        >= 1_024     => $"{bytesPerSec / 1_024.0:0.0} KB/s",
        _            => $"{bytesPerSec:0} B/s"
    };

    private static string FormatEta(TimeSpan eta)
    {
        if (eta.TotalHours   >= 1) return $"{(int)eta.TotalHours}h {eta.Minutes}m";
        if (eta.TotalMinutes >= 1) return $"{(int)eta.TotalMinutes}m {eta.Seconds}s";
        return $"{eta.Seconds}s";
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Sentinel exception for 401/403 responses
// ════════════════════════════════════════════════════════════════════════════

internal sealed class SessionExpiredException : Exception
{
    public SessionExpiredException() : base("Session expired — please log in again.") { }
}

// ════════════════════════════════════════════════════════════════════════════
// Data models
// ════════════════════════════════════════════════════════════════════════════

internal enum DownloadKind   { Video, Thumbnail }
internal enum SectionCategory { Drafts, Likes, Posts }
internal enum DownloadStatus  { Found, Downloading, Done, Failed }

/// <summary>
/// Observable item in the queue. Implements INotifyPropertyChanged so the DataGrid
/// updates live without manual Refresh() calls.
/// </summary>
internal sealed class DownloadItem : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    private DownloadKind    _kind;
    private SectionCategory _section;
    private string          _url         = "";
    private DownloadStatus  _status;
    private string          _progressText = "";
    private string?         _filePath;
    private string          _speed        = "";

    public DownloadKind    Kind         { get => _kind;         set => Set(ref _kind,         value); }
    public SectionCategory Section      { get => _section;      set => Set(ref _section,      value); }
    public string          Url          { get => _url;          set => Set(ref _url,          value); }
    public DownloadStatus  Status       { get => _status;       set => Set(ref _status,       value); }
    public string          ProgressText { get => _progressText; set => Set(ref _progressText, value); }
    public string?         FilePath     { get => _filePath;     set => Set(ref _filePath,     value); }
    public string          Speed        { get => _speed;        set => Set(ref _speed,        value); }
}

internal sealed class ScanResult
{
    public List<string> videoUrls { get; set; } = new();
    public List<string> thumbUrls { get; set; } = new();
    public List<string> pageLinks { get; set; } = new();
}

internal sealed class ProbeResult
{
    public string? url          { get; set; }
    public int     videos       { get; set; }
    public long    scrollHeight { get; set; }
}

internal sealed record PersistedItem(
    string  Section,
    string  Kind,
    string  Url,
    string  Status,
    string? FilePath);
