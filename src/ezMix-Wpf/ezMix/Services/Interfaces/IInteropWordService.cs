using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace ezMix.Services.Interfaces
{
    public interface IInteropWordService
    {
        // Mở tài liệu Word từ đường dẫn, có thể hiển thị hoặc ẩn
        Task<Document> OpenDocumentAsync(string filePath, bool visible);

        // Lưu tài liệu Word
        Task SaveDocumentAsync(_Document document);

        // Đóng một tài liệu Word
        Task CloseDocumentAsync(_Document document);

        // Đóng tất cả tài liệu Word đang mở
        Task CloseAllDocumentsAsync();

        // Thoát ứng dụng Word
        Task QuitWordAppAsync();

        // Định dạng lại toàn bộ tài liệu (font, style, layout...)
        Task FormatDocumentAsync(_Document document);

        // Thay thế nhiều chuỗi trong tài liệu theo danh sách
        Task ReplaceAsync(
            _Document document,
            Dictionary<string, string> replacements,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false);

        // Thay thế lần xuất hiện đầu tiên trong một đoạn văn
        Task ReplaceFirstAsync(
            Paragraph paragraph,
            string findText,
            string replaceWithText,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false);

        // Lặp lại việc thay thế cho đến khi hoàn tất hoặc đạt số lần tối đa
        Task ReplaceUntilDoneAsync(
            _Document document,
            Dictionary<string, string> replacements,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false,
            int maxIterations = 100);

        // Thay thế trong một section cụ thể của tài liệu
        Task ReplaceInSectionAsync(
            _Document document,
            int sectionIndex,
            string findText,
            string replaceWithText,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false);

        // Đổi chữ màu đỏ thành chữ gạch chân
        Task ReplaceRedTextWithUnderlineAsync(_Document document);

        // Đổi chữ gạch chân thành chữ màu đỏ
        Task ReplaceUnderlineWithRedTextAsync(_Document document);

        // Thay thế trong một khoảng ký tự (range) của tài liệu
        Task ReplaceInRangeAsync(
            _Document document,
            int start,
            int end,
            string findText,
            string replaceWithText,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false);

        // Chuyển danh sách (bullet/numbering) thành văn bản thường
        Task ConvertListFormatToTextAsync(_Document document);

        // Xóa tất cả header và footer trong tài liệu
        Task DeleteAllHeadersAndFootersAsync(_Document document);

        // Đặt đáp án thành A, B, C, D
        Task SetAnswersToABCDAsync(_Document document);

        // Đánh số thứ tự cho câu hỏi
        Task SetQuestionsToNumberAsync(_Document document);

        // Định dạng lại câu hỏi và câu trả lời
        Task FormatQuestionAndAnswerAsync(_Document document);

        // Cập nhật các trường (fields) trong tài liệu
        Task UpdateFieldsAsync(string filePath);

        // Xóa tất cả tab stop trong một đoạn văn
        Task ClearTabStopsAsync(Paragraph paragraph);

        // Sửa lỗi các công thức MathType trong tài liệu
        Task<int> FixMathTypeAsync(_Document document);

        // Chuyển đổi công thức sang định dạng MathType
        Task<int> ConvertEquationToMathTypeAsync(_Document document);

        // Từ chối tất cả thay đổi (track changes) trong tài liệu
        Task RejectAllChangesAsync(_Document document);

        void NormalizeParagraphEnds(_Document document);
    }
}
