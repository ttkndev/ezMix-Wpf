using ezMix.Core;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows;

namespace ezMix.ViewModels
{
    public class MainViewModel : ObservableObject
    {
        private readonly Version version = Assembly.GetExecutingAssembly().GetName().Version;

        private object currentView;
        public object CurrentView { get => currentView; set => SetProperty(ref currentView, value); }

        private bool isMenuExpanded = true;
        public bool IsMenuExpanded
        {
            get => isMenuExpanded;
            set
            {
                if (SetProperty(ref isMenuExpanded, value))
                {
                    OnPropertyChanged(nameof(VersionText));
                }
            }
        }

        private double menuWidth = 150;
        public double MenuWidth { get => menuWidth; set => SetProperty(ref menuWidth, value); }

        private string selectedMenuText = "Trang chủ";
        public string SelectedMenuText { get => selectedMenuText; set => SetProperty(ref selectedMenuText, value); }

        public string VersionText => IsMenuExpanded ? $"Phiên bản {version.Major}.{version.Minor}.{version.Build}" : $"{version.Major}.{version.Minor}.{version.Build}";

        public RelayCommand ChangeViewCommand { get; }
        public RelayCommand ToggleMenuCommand { get; }
        public RelayCommand OpenUrlCommand { get; }

        public MainViewModel()
        {
            ChangeViewCommand = new RelayCommand(ChangeView);
            ToggleMenuCommand = new RelayCommand(ToggleMenu);
            OpenUrlCommand = new RelayCommand(OpenUrl);

            CurrentView = App.ServiceProvider.GetRequiredService<HomeViewModel>();
        }

        private void ChangeView(object parameter)
        {
            switch (parameter?.ToString())
            {
                case "Home": CurrentView = App.ServiceProvider.GetRequiredService<HomeViewModel>(); SelectedMenuText = "Trang chủ"; break;
                case "Mix": CurrentView = App.ServiceProvider.GetRequiredService<MixViewModel>(); SelectedMenuText = "Trộn đề"; break;
                case "Utility": CurrentView = App.ServiceProvider.GetRequiredService<UtilityViewModel>(); SelectedMenuText = "Tiện ích"; break;
            }
        }

        private void ToggleMenu(object parameter)
        {
            IsMenuExpanded = !IsMenuExpanded;
            MenuWidth = IsMenuExpanded ? 150 : 50;
        }

        private void OpenUrl(object parameter)
        {
            string url = parameter?.ToString();
            if (!string.IsNullOrWhiteSpace(url))
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
        }
    }
}
