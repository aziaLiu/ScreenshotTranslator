using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;

namespace ScreenshotTranslator;

public partial class MainWindow : Window
{
    private readonly AppSettingsService _settingsService = new();
    private readonly MyMemoryTranslator _translator = new();
    private readonly ScreenshotOcrService _ocrService = new();
    private System.Windows.Forms.NotifyIcon? _notifyIcon;
    private GlobalHotkey? _hotkey;
    private AppSettings _settings = new();
    private HotkeyGesture _editingHotkey = HotkeyGesture.Default;
    private bool _isCapturing;
    private bool _isExitRequested;

    public MainWindow()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Closing += OnClosing;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _settings = _settingsService.Load();
        _editingHotkey = _settings.Hotkey;
        SourceLanguageComboBox.ItemsSource = OcrLanguageOption.Available;
        SourceLanguageComboBox.SelectedItem = OcrLanguageOption.FindByOcrTag(_settings.SourceLanguageTag);
        UpdateHotkeyDisplay(_settings.Hotkey);
        CreateTrayIcon();
        RegisterHotkey(_settings.Hotkey);
    }

    private void OnClosing(object? sender, CancelEventArgs e)
    {
        if (!_isExitRequested)
        {
            e.Cancel = true;
            Hide();
            SetStatus("已最小化到托盘，热键仍然可用。");
            return;
        }

        _hotkey?.Dispose();
        _notifyIcon?.Dispose();
    }

    protected override void OnStateChanged(EventArgs e)
    {
        base.OnStateChanged(e);

        if (WindowState == WindowState.Minimized)
        {
            Hide();
            SetStatus("已最小化到托盘，热键仍然可用。");
        }
    }

    private void CreateTrayIcon()
    {
        if (_notifyIcon is not null)
        {
            return;
        }

        var menu = new System.Windows.Forms.ContextMenuStrip();
        menu.Items.Add("显示窗口", null, (_, _) => ShowMainWindow());
        menu.Items.Add("退出", null, (_, _) =>
        {
            _isExitRequested = true;
            Close();
        });

        _notifyIcon = new System.Windows.Forms.NotifyIcon
        {
            Icon = LoadTrayIcon(),
            Text = "截图翻译",
            Visible = true,
            ContextMenuStrip = menu
        };

        _notifyIcon.DoubleClick += (_, _) => ShowMainWindow();
    }

    private static System.Drawing.Icon LoadTrayIcon()
    {
        var iconResource = System.Windows.Application.GetResourceStream(new Uri("pack://application:,,,/Assets/AppIcon.ico"));
        if (iconResource?.Stream is not null)
        {
            return new System.Drawing.Icon(iconResource.Stream);
        }

        return System.Drawing.SystemIcons.Application;
    }

    public void ShowMainWindow()
    {
        Show();
        WindowState = WindowState.Normal;
        Activate();
        Topmost = true;
        Topmost = false;
        Focus();
    }

    private void RegisterHotkey(HotkeyGesture gesture)
    {
        _hotkey?.Dispose();
        var handle = new WindowInteropHelper(this).Handle;
        _hotkey = new GlobalHotkey(handle, gesture);
        _hotkey.Pressed += async (_, _) => await CaptureTranslateAsync();

        if (_hotkey.Register())
        {
            SetStatus($"热键已启用：{gesture}");
            return;
        }

        SetStatus($"热键注册失败，可能已被占用：{gesture}");
    }

    private async void CaptureButton_OnClick(object sender, RoutedEventArgs e)
    {
        await CaptureTranslateAsync();
    }

    private async void TranslateEditedTextButton_OnClick(object sender, RoutedEventArgs e)
    {
        await TranslateCurrentTextAsync();
    }

    private void SaveHotkeyButton_OnClick(object sender, RoutedEventArgs e)
    {
        _settings.Hotkey = _editingHotkey;
        _settings.SourceLanguageTag = GetSelectedLanguage().OcrLanguageTag;
        _settingsService.Save(_settings);
        UpdateHotkeyDisplay(_settings.Hotkey);
        RegisterHotkey(_settings.Hotkey);
    }

    private void SourceLanguageComboBox_OnSelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (SourceLanguageComboBox.SelectedItem is not OcrLanguageOption option)
        {
            return;
        }

        _settings.SourceLanguageTag = option.OcrLanguageTag;
        _settingsService.Save(_settings);
        SetStatus($"识别语言已切换：{option.DisplayName}");
    }

    private void HotkeyTextBox_OnGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        HotkeyTextBox.Text = "请按下组合键...";
    }

    private void HotkeyTextBox_OnPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        e.Handled = true;

        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        if (key is Key.LeftCtrl or Key.RightCtrl or Key.LeftAlt or Key.RightAlt or Key.LeftShift or Key.RightShift or Key.LWin or Key.RWin)
        {
            return;
        }

        var modifiers = Keyboard.Modifiers;
        if (modifiers == ModifierKeys.None)
        {
            SetStatus("请至少使用一个修饰键，例如 Ctrl+Shift+T。");
            return;
        }

        _editingHotkey = new HotkeyGesture(modifiers, key);
        HotkeyTextBox.Text = _editingHotkey.ToString();
    }

    private async Task CaptureTranslateAsync()
    {
        if (_isCapturing)
        {
            return;
        }

        _isCapturing = true;
        CaptureButton.IsEnabled = false;

        try
        {
            OcrTextBox.Text = "";
            TranslationTextBox.Text = "";
            SetStatus("请框选要翻译的屏幕区域...");

            var image = await CaptureWithSystemSnippingToolAsync(TimeSpan.FromSeconds(45));
            if (image is null)
            {
                SetStatus("未检测到截图。");
                return;
            }

            var sourceLanguage = GetSelectedLanguage();
            SetStatus($"正在使用 {sourceLanguage.DisplayName} OCR 识别文字...");
            var recognizedText = await _ocrService.RecognizeAsync(image, sourceLanguage.OcrLanguageTag);
            OcrTextBox.Text = recognizedText;

            if (string.IsNullOrWhiteSpace(recognizedText))
            {
                SetStatus("未识别到文字。");
                return;
            }

            await TranslateCurrentTextAsync();
            ShowMainWindow();
        }
        catch (Exception ex)
        {
            SetStatus($"失败：{ex.Message}");
        }
        finally
        {
            CaptureButton.IsEnabled = true;
            _isCapturing = false;
        }
    }

    private async Task TranslateCurrentTextAsync()
    {
        var text = OcrTextBox.Text;
        if (string.IsNullOrWhiteSpace(text))
        {
            SetStatus("识别文本为空，无法翻译。");
            return;
        }

        CaptureButton.IsEnabled = false;
        TranslateEditedTextButton.IsEnabled = false;

        try
        {
            var sourceLanguage = GetSelectedLanguage();
            SetStatus("正在翻译为中文...");
            TranslationTextBox.Text = await _translator.TranslateToChineseAsync(text, sourceLanguage.TranslationSourceCode);
            SetStatus("完成");
        }
        catch (Exception ex)
        {
            SetStatus($"翻译失败：{ex.Message}");
        }
        finally
        {
            CaptureButton.IsEnabled = true;
            TranslateEditedTextButton.IsEnabled = true;
        }
    }

    private async Task<BitmapSource?> CaptureWithSystemSnippingToolAsync(TimeSpan timeout)
    {
        var before = NativeMethods.GetClipboardSequenceNumber();
        Process.Start(new ProcessStartInfo("ms-screenclip:") { UseShellExecute = true });

        var start = DateTimeOffset.UtcNow;
        while (DateTimeOffset.UtcNow - start < timeout)
        {
            await Task.Delay(350);
            if (NativeMethods.GetClipboardSequenceNumber() == before || !System.Windows.Clipboard.ContainsImage())
            {
                continue;
            }

            var bitmap = System.Windows.Clipboard.GetImage();
            bitmap?.Freeze();
            return bitmap;
        }

        return null;
    }

    private void UpdateHotkeyDisplay(HotkeyGesture gesture)
    {
        var text = gesture.ToString();
        HotkeyTextBox.Text = text;
        HotkeyBadge.Text = text;
    }

    private void SetStatus(string message)
    {
        StatusTextBlock.Text = message;
    }

    private OcrLanguageOption GetSelectedLanguage()
    {
        return SourceLanguageComboBox.SelectedItem as OcrLanguageOption ?? OcrLanguageOption.Default;
    }
}

public sealed class ScreenshotOcrService
{
    public async Task<string> RecognizeAsync(BitmapSource bitmapSource, string languageTag)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"screenshot-translator-{Guid.NewGuid():N}.png");

        try
        {
            var preparedBitmap = PrepareForOcr(bitmapSource);
            SaveBitmap(preparedBitmap, tempPath);
            var file = await StorageFile.GetFileFromPathAsync(tempPath);
            using var stream = await file.OpenAsync(FileAccessMode.Read);
            var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream);
            using var softwareBitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);

            var language = new Language(languageTag);
            var engine = OcrEngine.IsLanguageSupported(language)
                ? OcrEngine.TryCreateFromLanguage(language)
                : OcrEngine.TryCreateFromUserProfileLanguages();

            if (engine is null)
            {
                throw new InvalidOperationException("当前系统没有可用的 Windows OCR 语言包。");
            }

            var result = await engine.RecognizeAsync(softwareBitmap);
            return result.Text.Trim();
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    private static BitmapSource PrepareForOcr(BitmapSource source)
    {
        var shortestSide = Math.Min(source.PixelWidth, source.PixelHeight);
        var longestSide = Math.Max(source.PixelWidth, source.PixelHeight);
        var scaleForSmallText = shortestSide < 900 ? 900d / Math.Max(shortestSide, 1) : 1d;
        var scale = Math.Clamp(scaleForSmallText, 1d, 3d);

        if (longestSide * scale > OcrEngine.MaxImageDimension)
        {
            scale = OcrEngine.MaxImageDimension / (double)longestSide;
        }

        var transformed = scale > 1.01d
            ? new TransformedBitmap(source, new ScaleTransform(scale, scale))
            : source;

        transformed.Freeze();
        return transformed;
    }

    private static void SaveBitmap(BitmapSource bitmapSource, string path)
    {
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmapSource));

        using var stream = File.Create(path);
        encoder.Save(stream);
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }
}

public sealed class MyMemoryTranslator
{
    private const int MaxQueryLength = 480;

    private static readonly HttpClient Client = new()
    {
        Timeout = TimeSpan.FromSeconds(20)
    };

    public async Task<string> TranslateToChineseAsync(string text, string sourceLanguageCode)
    {
        var chunks = SplitForMyMemory(text);
        var translatedChunks = new List<string>(chunks.Count);

        foreach (var chunk in chunks)
        {
            translatedChunks.Add(await TranslateChunkToChineseAsync(chunk, sourceLanguageCode));
            await Task.Delay(150);
        }

        return string.Join(Environment.NewLine, translatedChunks.Where(chunk => !string.IsNullOrWhiteSpace(chunk)));
    }

    private static List<string> SplitForMyMemory(string text)
    {
        var chunks = new List<string>();
        var paragraphs = text.Replace("\r\n", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var paragraph in paragraphs)
        {
            SplitParagraph(paragraph, chunks);
        }

        return chunks;
    }

    private static void SplitParagraph(string paragraph, List<string> chunks)
    {
        var current = "";
        var sentences = paragraph.Split(['.', '!', '?', ';', '。', '！', '？', '；'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var sentence in sentences)
        {
            if (sentence.Length > MaxQueryLength)
            {
                FlushCurrent(chunks, ref current);
                SplitLongText(sentence, chunks);
                continue;
            }

            if (current.Length + sentence.Length + 1 > MaxQueryLength)
            {
                FlushCurrent(chunks, ref current);
            }

            current = string.IsNullOrWhiteSpace(current) ? sentence : $"{current} {sentence}";
        }

        FlushCurrent(chunks, ref current);
    }

    private static void SplitLongText(string text, List<string> chunks)
    {
        for (var index = 0; index < text.Length; index += MaxQueryLength)
        {
            var length = Math.Min(MaxQueryLength, text.Length - index);
            chunks.Add(text.Substring(index, length));
        }
    }

    private static void FlushCurrent(List<string> chunks, ref string current)
    {
        if (!string.IsNullOrWhiteSpace(current))
        {
            chunks.Add(current.Trim());
            current = "";
        }
    }

    private async Task<string> TranslateChunkToChineseAsync(string text, string sourceLanguageCode)
    {
        var url = "https://api.mymemory.translated.net/get?q=" +
                  Uri.EscapeDataString(text) +
                  $"&langpair={Uri.EscapeDataString(sourceLanguageCode)}%7Czh-CN";

        using var response = await Client.GetAsync(url);
        response.EnsureSuccessStatusCode();

        var payload = await response.Content.ReadFromJsonAsync<MyMemoryResponse>();
        if (payload?.ResponseStatus is >= 400)
        {
            return $"翻译失败：{payload.ResponseDetails ?? "翻译接口返回错误。"}";
        }

        var translatedText = payload?.ResponseData?.TranslatedText;
        return string.IsNullOrWhiteSpace(translatedText) ? "翻译接口未返回结果。" : translatedText;
    }

    private sealed class MyMemoryResponse
    {
        public MyMemoryResponseData? ResponseData { get; set; }

        public int ResponseStatus { get; set; }

        public string? ResponseDetails { get; set; }
    }

    private sealed class MyMemoryResponseData
    {
        public string? TranslatedText { get; set; }
    }
}

public sealed class AppSettingsService
{
    private readonly string _settingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "ScreenshotTranslator",
        "settings.json");

    public AppSettings Load()
    {
        try
        {
            if (!File.Exists(_settingsPath))
            {
                return new AppSettings();
            }

            var json = File.ReadAllText(_settingsPath);
            return System.Text.Json.JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_settingsPath)!);
        var json = System.Text.Json.JsonSerializer.Serialize(settings, new System.Text.Json.JsonSerializerOptions
        {
            WriteIndented = true
        });
        File.WriteAllText(_settingsPath, json);
    }
}

public sealed class AppSettings
{
    public HotkeyGesture Hotkey { get; set; } = HotkeyGesture.Default;

    public string SourceLanguageTag { get; set; } = OcrLanguageOption.Default.OcrLanguageTag;
}

public sealed record OcrLanguageOption(string DisplayName, string OcrLanguageTag, string TranslationSourceCode)
{
    public static OcrLanguageOption Default { get; } = new("英文", "en-US", "en");

    public static IReadOnlyList<OcrLanguageOption> Available { get; } =
    [
        Default,
        new("日文", "ja-JP", "ja"),
        new("韩文", "ko-KR", "ko"),
        new("简体中文", "zh-CN", "zh-CN"),
        new("繁体中文", "zh-TW", "zh-TW"),
        new("法文", "fr-FR", "fr"),
        new("德文", "de-DE", "de"),
        new("西班牙文", "es-ES", "es"),
        new("俄文", "ru-RU", "ru")
    ];

    public static OcrLanguageOption FindByOcrTag(string? tag)
    {
        return Available.FirstOrDefault(option => string.Equals(option.OcrLanguageTag, tag, StringComparison.OrdinalIgnoreCase)) ?? Default;
    }
}

public readonly record struct HotkeyGesture(ModifierKeys Modifiers, Key Key)
{
    public static HotkeyGesture Default { get; } = new(ModifierKeys.Control | ModifierKeys.Shift, Key.T);

    public override string ToString()
    {
        var parts = new List<string>();
        if (Modifiers.HasFlag(ModifierKeys.Control))
        {
            parts.Add("Ctrl");
        }

        if (Modifiers.HasFlag(ModifierKeys.Alt))
        {
            parts.Add("Alt");
        }

        if (Modifiers.HasFlag(ModifierKeys.Shift))
        {
            parts.Add("Shift");
        }

        if (Modifiers.HasFlag(ModifierKeys.Windows))
        {
            parts.Add("Win");
        }

        parts.Add(Key.ToString());
        return string.Join("+", parts);
    }
}

public sealed class GlobalHotkey : IDisposable
{
    private const int HotkeyId = 0x4A21;
    private const int WmHotkey = 0x0312;
    private readonly nint _windowHandle;
    private readonly HotkeyGesture _gesture;
    private HwndSource? _source;
    private bool _registered;

    public event EventHandler? Pressed;

    public GlobalHotkey(nint windowHandle, HotkeyGesture gesture)
    {
        _windowHandle = windowHandle;
        _gesture = gesture;
    }

    public bool Register()
    {
        _source = HwndSource.FromHwnd(_windowHandle);
        _source?.AddHook(WndProc);

        var modifiers = ToNativeModifiers(_gesture.Modifiers);
        var virtualKey = KeyInterop.VirtualKeyFromKey(_gesture.Key);
        _registered = NativeMethods.RegisterHotKey(_windowHandle, HotkeyId, modifiers, virtualKey);
        return _registered;
    }

    public void Dispose()
    {
        if (_registered)
        {
            NativeMethods.UnregisterHotKey(_windowHandle, HotkeyId);
            _registered = false;
        }

        _source?.RemoveHook(WndProc);
        _source = null;
    }

    private nint WndProc(nint hwnd, int msg, nint wParam, nint lParam, ref bool handled)
    {
        if (msg == WmHotkey && wParam.ToInt32() == HotkeyId)
        {
            handled = true;
            Pressed?.Invoke(this, EventArgs.Empty);
        }

        return nint.Zero;
    }

    private static uint ToNativeModifiers(ModifierKeys modifiers)
    {
        uint value = 0;
        if (modifiers.HasFlag(ModifierKeys.Alt))
        {
            value |= 0x0001;
        }

        if (modifiers.HasFlag(ModifierKeys.Control))
        {
            value |= 0x0002;
        }

        if (modifiers.HasFlag(ModifierKeys.Shift))
        {
            value |= 0x0004;
        }

        if (modifiers.HasFlag(ModifierKeys.Windows))
        {
            value |= 0x0008;
        }

        return value;
    }
}

internal static class NativeMethods
{
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool RegisterHotKey(nint hWnd, int id, uint fsModifiers, int vk);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UnregisterHotKey(nint hWnd, int id);

    [DllImport("user32.dll")]
    public static extern uint GetClipboardSequenceNumber();
}
