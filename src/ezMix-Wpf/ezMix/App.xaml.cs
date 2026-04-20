using ezMix.Helpers;
using ezMix.Services.Implementations;
using ezMix.Services.Interfaces;
using ezMix.ViewModels;
using ezMix.Views;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

namespace ezMix
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static ServiceProvider ServiceProvider { get; private set; } = null;

        public App()
        {
            var services = new ServiceCollection();

            services.AddTransient<IExcelService, ExcelService>();
            services.AddTransient<IOpenXMLService, OpenXMLService>();
            services.AddTransient<IInteropWordService, InteropWordService>();
            services.AddTransient<IGeminiService, GeminiService>();
            services.AddSingleton<IUpdateCheckService, UpdateCheckService>();
            services.AddSingleton<HttpClient>();

            services.AddSingleton<MainViewModel>();
            services.AddTransient<HomeViewModel>();
            services.AddTransient<MixViewModel>();
            services.AddTransient<UtilityViewModel>();

            ServiceProvider = services.BuildServiceProvider();
        }

        protected override async void OnStartup(StartupEventArgs e)
        {      
            base.OnStartup(e);

            bool updaterStarted = false;
            if (await InternetHelper.IsInternetAvailableAsync())
            {
                updaterStarted = await TryCheckAndStartUpdateAsync();
            }

            if (!updaterStarted)
            {
                var mainWindow = new MainWindow
                {
                    DataContext = ServiceProvider.GetRequiredService<MainViewModel>()
                }; 
                mainWindow.Show();
            }
        }

        private async Task<bool> TryCheckAndStartUpdateAsync()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            string currentVersion = $"{version.Major}.{version.Minor}.{version.Build}";

            const string updateJsonUrl = "https://raw.githubusercontent.com/nhathinh2703/ez-updates/main/apps/ezMix/latest.json";

            var updateService = ServiceProvider.GetRequiredService<IUpdateCheckService>();
            try
            {
                var latest = await updateService.GetLatestAsync(updateJsonUrl);
                if (latest != null && updateService.HasUpdate(currentVersion, latest.Version))
                {
                    var message =
                        $"Phiên bản hiện tại: {currentVersion}\n" +
                        $"Có phiên bản mới: {latest.Version}\n\n" +
                        $"{string.Join("\n", latest.Changelog)}\n\n" +
                        "Cập nhật ngay?";

                    var result = MessageBox.Show(
                        message,
                        "Cập nhật phần mềm",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information);

                    if (result == MessageBoxResult.Yes || latest.Mandatory)
                    {
                        StartUpdater(latest.Url);
                        return true;   // ✅ ĐÃ BẮT ĐẦU UPDATE
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi khi kiểm tra update: {ex.Message}");
                // Không làm app chết nếu lỗi mạng / json
            }

            return false; // ❌ Không update → tiếp tục chạy app
        }

        private void StartUpdater(string zipUrl)
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string updaterExe = Path.Combine(appDir, "Updater", "ezUpdater.Net48.exe");

            if (!File.Exists(updaterExe))
                return;

            int pid = Process.GetCurrentProcess().Id;

            string args = string.Format(
                "--app-dir \"{0}\" --zip-url \"{1}\" --exe-name \"{2}\" --parent-pid {3}",
                appDir.TrimEnd('\\'),   // bỏ dấu \ cuối
                zipUrl,
                "ezMix.exe",
                pid);
            Debug.WriteLine("Updater args: " + args);

            var psi = new ProcessStartInfo
            {
                FileName = updaterExe,
                Arguments = args,
                UseShellExecute = false,   // để arguments truyền đúng
                CreateNoWindow = false
            };

            Process.Start(psi);
            Shutdown();
        }
    }
}
