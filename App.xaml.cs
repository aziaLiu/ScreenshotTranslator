using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Windows;

namespace ScreenshotTranslator;

public partial class App : System.Windows.Application
{
    private const string SingleInstanceMutexName = @"Local\ScreenshotTranslator.SingleInstance";
    private const string ShowWindowPipeName = "ScreenshotTranslator.ShowWindow";

    private Mutex? _singleInstanceMutex;
    private CancellationTokenSource? _pipeCancellation;
    private MainWindow? _mainWindow;
    private bool _ownsSingleInstanceMutex;

    protected override async void OnStartup(StartupEventArgs e)
    {
        _singleInstanceMutex = new Mutex(false, SingleInstanceMutexName);
        _ownsSingleInstanceMutex = _singleInstanceMutex.WaitOne(0);

        if (!_ownsSingleInstanceMutex)
        {
            await SignalExistingInstanceAsync();
            Shutdown();
            Environment.Exit(0);
            return;
        }

        base.OnStartup(e);

        _pipeCancellation = new CancellationTokenSource();
        _ = ListenForShowRequestsAsync(_pipeCancellation.Token);

        _mainWindow = new MainWindow();
        MainWindow = _mainWindow;
        _mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _pipeCancellation?.Cancel();
        _pipeCancellation?.Dispose();
        if (_ownsSingleInstanceMutex)
        {
            _singleInstanceMutex?.ReleaseMutex();
        }

        _singleInstanceMutex?.Dispose();
        base.OnExit(e);
    }

    private static async Task SignalExistingInstanceAsync()
    {
        try
        {
            using var client = new NamedPipeClientStream(".", ShowWindowPipeName, PipeDirection.Out);
            await client.ConnectAsync(1200);
            await using var writer = new StreamWriter(client);
            await writer.WriteLineAsync("show");
            await writer.FlushAsync();
        }
        catch
        {
            // The first instance may still be starting; avoid opening a second UI.
        }
    }

    private async Task ListenForShowRequestsAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await using var server = new NamedPipeServerStream(
                    ShowWindowPipeName,
                    PipeDirection.In,
                    1,
                    PipeTransmissionMode.Message,
                    PipeOptions.Asynchronous);

                await server.WaitForConnectionAsync(cancellationToken);
                using var reader = new StreamReader(server);
                var message = await reader.ReadLineAsync(cancellationToken);

                if (string.Equals(message, "show", StringComparison.OrdinalIgnoreCase))
                {
                    Dispatcher.Invoke(() => _mainWindow?.ShowMainWindow());
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
            }
        }
    }
}
