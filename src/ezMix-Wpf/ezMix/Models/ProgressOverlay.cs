using ezMix.Core;

namespace ezMix.Models
{
    public class ProgressOverlay : ObservableObject
    {
        private string title;
        private bool isVisible = false;
        private bool isIndeterminate = false;
        private double progressValue = 0;
        private string statusText = "Đang xử lý...";

        public bool IsVisible { get => isVisible; set => SetProperty(ref isVisible, value); }
        public bool IsIndeterminate { get => isIndeterminate; set => SetProperty(ref isIndeterminate, value); }
        public double ProgressValue { get => progressValue; set => SetProperty(ref progressValue, value); }
        public string StatusText { get => statusText; set => SetProperty(ref statusText, value); }
        public string Title { get => title; set => SetProperty(ref title, value); }
    }
}
