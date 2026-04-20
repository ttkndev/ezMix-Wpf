using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ezMix.Assets.Converters
{
    public class WidthToMarginConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double width = (double)value;
            // Đặt nút ngay sát biên phải của sidebar
            return new Thickness(width - 15, 0, 0, 0); // 15 = nửa chiều rộng nút
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

}
