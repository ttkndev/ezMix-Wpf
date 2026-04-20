using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ezMix.Core
{
    public class ObservableObject : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gọi sự kiện PropertyChanged để thông báo cho UI cập nhật. 
        /// </summary> 
        /// <param name="propertyName">Tên property (tự động lấy nếu không truyền)</param> 
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Hàm SetProperty giúp giảm lặp code khi gán giá trị cho property. 
        /// </summary> 
        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (Equals(field, value))
                return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
}
