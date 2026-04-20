using System;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace ezMix.Helpers
{
    public class MessageHelper
    {
        // ✅ Hiển thị thông báo thành công
        public static MessageBoxResult Success(string message, string title = "Thông báo", MessageBoxImage icon = MessageBoxImage.Information)
        {
            // Hiển thị hộp thoại với nút OK và biểu tượng "Thông tin"
            return MessageBox.Show(message, title, MessageBoxButton.OK, icon);
        }

        // ❓ Hiển thị câu hỏi xác nhận
        public static MessageBoxResult Question(string message, string title = "Xác nhận", MessageBoxImage icon = MessageBoxImage.Question)
        {
            // Hiển thị hộp thoại với nút Yes/No và biểu tượng "Câu hỏi"
            return MessageBox.Show(message, title, MessageBoxButton.YesNo, icon);
        }

        // ❌ Hiển thị thông báo lỗi với chuỗi lỗi
        public static MessageBoxResult Error(string message, string title = "Lỗi", MessageBoxImage icon = MessageBoxImage.Error)
        {
            // Thêm tiền tố "Thất bại!" và nội dung lỗi
            return MessageBox.Show("Thất bại!\nLỗi: " + message, title, MessageBoxButton.OK, icon);
        }

        // ❌ Hiển thị thông báo lỗi với đối tượng Exception
        public static MessageBoxResult Error(Exception ex, string title = "Lỗi", MessageBoxImage icon = MessageBoxImage.Error)
        {
            // Lấy thông tin lỗi từ Exception.Message
            return MessageBox.Show("Thất bại!\nLỗi: " + ex.Message, title, MessageBoxButton.OK, icon);
        }

    }
}
