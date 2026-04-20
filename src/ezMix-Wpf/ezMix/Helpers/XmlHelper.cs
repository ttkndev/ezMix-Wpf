using System.IO;
using System.Xml.Serialization;

namespace ezMix.Helpers
{
    public class XmlHelper
    {
        public static T LoadFromXml<T>(string filePath) where T : new()
        {
            // 📂 Nếu file không tồn tại → trả về đối tượng mới mặc định (new T)
            if (!File.Exists(filePath)) return new T();

            // 📖 Mở file ở chế độ đọc
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                // 🛠️ Tạo đối tượng XmlSerializer để chuyển đổi dữ liệu XML thành đối tượng kiểu T
                var serializer = new XmlSerializer(typeof(T));

                // 🔄 Deserialize: đọc dữ liệu XML và ánh xạ thành đối tượng T
                return (T)serializer.Deserialize(stream);
            }
        }

        public static void SaveToXml<T>(string filePath, T data)
        {
            // 📖 Mở (hoặc tạo mới) file ở chế độ ghi
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                // 🛠️ Tạo đối tượng XmlSerializer để chuyển đổi dữ liệu đối tượng thành XML
                var serializer = new XmlSerializer(typeof(T));

                // 🔄 Serialize: ghi dữ liệu đối tượng T thành XML vào file
                serializer.Serialize(stream, data);
            }
        }
    }
}
