namespace ezMix.Models
{
    public class MixInfo
    {
        public string Code { get; set; } = string.Empty;    // Mã đề
        public int NumberOfVersions { get; set; } = 4;      // Số đề cần trộn
        public string[] Versions { get; set; } = new string[0];        // Danh sách mã đề
        public string StartCode { get; set; } = "01";       // Mã đề bắt đầu
        public string SuperiorUnit { get; set; } = "SỞ GDĐT ...";   // Đơn vị cấp trên
        public string Unit { get; set; } = "TRƯỜNG THPT ...";       // Đơn vị
        public string TestPeriod { get; set; } = "ĐỀ KIỂM TRA GIỮA KỲ 1";      // Tên bài kiểm tra
        public string Grade { get; set; } = "12";                           // Khối lớp
        public string SchoolYear { get; set; } = "2025-2026";               // Năm học
        public string Subject { get; set; } = "TIN HỌC";                    // Môn học
        public string Time { get; set; } = "45 phút";                       // Thời gian làm bài

        public string FontFamily { get; set; } = "Times New Roman";         // Phông chữ
        public string FontSize { get; set; } = "12";                        // Cỡ chữ

        public bool IsFixMathType { get; set; } = false;             // Sửa công thức MathType

        public bool IsShuffledQuestionMultipleChoice { get; set; } = true;
        public bool IsShuffledAnswerMultipleChoice { get; set; } = true;
        public bool IsShuffledQuestionTrueFalse { get; set; } = true;
        public bool IsShuffledAnswerTrueFalse { get; set; } = true;
        public bool IsShuffledShortAnswer { get; set; } = true;
        public bool IsShuffledEssay { get; set; } = true;
        public bool IsShowWordWhenAnalyze { get; set; } = true;

        public string PointMultipleChoice { get; set; } = "3,0";
        public string PointTrueFalse { get; set; } = "2,0";
        public string PointShortAnswer { get; set; } = "2,0";
        public string PointEssay { get; set; } = "3,0";


        public string GeminiApiKey { get; set; } = "Nhập key của bạn";
        public string GeminiModel { get; set; } = "gemini-2.5-flash";
    }
}
