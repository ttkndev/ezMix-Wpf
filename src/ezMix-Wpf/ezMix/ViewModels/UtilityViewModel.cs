using ezMix.Core;
using ezMix.Helpers;
using ezMix.Services.Interfaces;
using Regex.Helpers;
using System;
using System.Threading.Tasks;

namespace ezMix.ViewModels
{
    public class UtilityViewModel : ObservableObject
    {
        private readonly IInteropWordService _interopWordService;

        public RelayCommand FixMathTypeCommand { get; }
        public RelayCommand OpenResourceCommand { get; }

        public UtilityViewModel(IInteropWordService interopWordService)
        {
            _interopWordService = interopWordService;

            FixMathTypeCommand = new RelayCommand(async _ => await FixMathType());
            OpenResourceCommand = new RelayCommand(OpenResource);
        }

        private async Task FixMathType()
        {
            try
            {
                var filePath = FileHelper.BrowseFile();
                if (string.IsNullOrEmpty(filePath))
                    return;

                var document = await _interopWordService.OpenDocumentAsync(filePath, visible: true);
                document.Activate();

                int count = await _interopWordService.FixMathTypeAsync(document);

                MessageHelper.Success($"✅ Đã xử lý {count} công thức MathType.");
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        private void OpenResource(object parameter)
        {
            string url = parameter?.ToString();
            if (!string.IsNullOrWhiteSpace(url))
            {
                FileHelper.OpenResource(url);
            }
        }
    }
}
