using System.ComponentModel;

namespace ezMix.Models.Enums
{
    public enum QuestionType
    {
        [Description("TN Nhiều lựa chọn")]
        MultipleChoice,
        [Description("TN Đúng/sai")]
        TrueFalse,
        [Description("TN Trả lời ngắn")]
        ShortAnswer,
        [Description("Tự luận")]
        Essay,
        [Description("Chưa xác định")]
        Unknown
    }
}
