using System.Threading;
using System.Windows;

namespace SoraBulkDownloader.App;

public partial class App : System.Windows.Application
{
    private static Mutex? _mutex;

    protected override void OnStartup(StartupEventArgs e)
    {
        try
        {
            _mutex = new Mutex(true, "SoraBulkDownloader_SingleInstance", out var createdNew);
            if (!createdNew)
            {
                System.Windows.MessageBox.Show("Sora Bulk Downloader is already running.",
                    "Already Running", MessageBoxButton.OK, MessageBoxImage.Information);
                _mutex = null;
                Shutdown();
                return;
            }

            DispatcherUnhandledException += (_, args) =>
            {
                System.Windows.MessageBox.Show($"Unexpected error:\n\n{args.Exception.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                args.Handled = true;
            };

            AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            {
                var msg = (args.ExceptionObject as Exception)?.Message ?? args.ExceptionObject.ToString();
                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "sora_crash.txt"),
                    msg);
                System.Windows.MessageBox.Show($"Fatal error:\n\n{msg}", "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
            };

            base.OnStartup(e);
        }
        catch (Exception ex)
        {
            System.IO.File.WriteAllText(
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "sora_crash.txt"),
                ex.ToString());
            System.Windows.MessageBox.Show($"Startup error:\n\n{ex.Message}", "Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _mutex?.ReleaseMutex();
        _mutex?.Dispose();
        base.OnExit(e);
    }
}

