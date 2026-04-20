using ezMix.Helpers;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;

namespace Regex.Helpers
{
    public static class FileHelper
    {
        public static string BrowseFile(string title = "Chọn file Word", string filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*")
        {
            var dialog = new OpenFileDialog
            {
                Title = title,          // Tiêu đề hộp thoại
                Filter = filter,        // Bộ lọc loại file (Word, PDF, ảnh,...)
                CheckFileExists = true, // Kiểm tra file có tồn tại không
                //InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) // Thư mục mặc định
            };

            // Nếu người dùng chọn file → trả về đường dẫn, ngược lại trả về null
            return dialog.ShowDialog() == true ? dialog.FileName : null;
        }

        public static string BrowsePdf()
        {
            return BrowseFile("Chọn tệp PDF", "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*");
        }

        public static string BrowseImage()
        {
            return BrowseFile("Chọn tệp ảnh", "Image Files (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|All Files (*.*)|*.*");
        }

        public static void OpenResource(string pathOrUrl)
        {
            if (string.IsNullOrWhiteSpace(pathOrUrl))
            {
                MessageHelper.Error("Đường dẫn hoặc URL rỗng");
                return;
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = pathOrUrl,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Không thể mở resource: {ex.Message}");
            }
        }
    }
}
