using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ezMix.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ezMix.Services.Interfaces
{
    public interface IOpenXMLService
    {
        // Tạo mới một tài liệu Word tại đường dẫn cho trước
        Task<WordprocessingDocument> CreateDocumentAsync(string filePath);

        // Mở tài liệu Word từ đường dẫn, có thể chỉnh sửa bảng nếu cần
        Task<WordprocessingDocument> OpenDocumentAsync(string filePath, bool isEditTable);

        // Lưu tài liệu Word
        Task SaveDocumentAsync(WordprocessingDocument document);

        // Đóng tài liệu Word
        Task CloseDocumentAsync(WordprocessingDocument document);

        // Lấy toàn bộ nội dung văn bản từ tài liệu
        Task<string> GetDocumentTextAsync(WordprocessingDocument document);

        // Lấy phần thân (Body) của tài liệu
        Task<Body> GetDocumentBodyAsync(WordprocessingDocument document);

        // Định dạng tất cả các đoạn văn trong tài liệu theo thông tin mixInfo
        Task FormatAllParagraphsAsync(WordprocessingDocument doc, MixInfo mixInfo);

        // Phân tích file docx để lấy danh sách câu hỏi
        Task<List<Question>> ParseDocxQuestionsAsync(string filePath);

        // Trích xuất toàn bộ văn bản từ file docx
        Task<string> ExtractTextAsync(string filePath);

        /// <summary>
        /// Shuffle câu hỏi trong tài liệu Word theo phiên bản và thông tin MixInfo.
        /// </summary>
        /// <param name="doc">Tài liệu đề gốc (WordprocessingDocument).</param>
        /// <param name="version">Mã phiên bản đề.</param>
        /// <param name="answerDoc">Tài liệu đáp án (nếu có).</param>
        /// <param name="mixInfo">Thông tin cấu hình trộn đề.</param>
        /// <returns>Danh sách câu hỏi sau khi shuffle.</returns>
        Task<List<QuestionExport>> ShuffleQuestionsAsync(
            WordprocessingDocument doc,
            string version,
            WordprocessingDocument answerDoc = null,
            MixInfo mixInfo = null
        );

        /// <summary>
        /// Chèn template vào tài liệu Word theo thông tin MixInfo và mã đề.
        /// </summary>
        /// <param name="templatePath">Đường dẫn tới file template.</param>
        /// <param name="doc">Tài liệu Word cần chèn template.</param>
        /// <param name="mixInfo">Thông tin cấu hình trộn đề.</param>
        /// <param name="code">Mã đề kiểm tra.</param>
        Task InsertTemplateAsync(
            string templatePath,
            WordprocessingDocument doc,
            MixInfo mixInfo,
            string code
        );

        /// <summary>
        /// Thêm EndNotes vào tài liệu Word.
        /// </summary>
        /// <param name="doc">Tài liệu Word cần thêm EndNotes.</param>
        Task AddEndNotesAsync(WordprocessingDocument doc);

        /// <summary>
        /// Chèn phần hướng dẫn (guide) vào tài liệu Word dựa trên danh sách đáp án, thông tin MixInfo và mã đề.
        /// </summary>
        /// <param name="doc">Tài liệu Word cần chèn guide.</param>
        /// <param name="answers">Danh sách đáp án của đề.</param>
        /// <param name="mixInfo">Thông tin cấu hình trộn đề.</param>
        /// <param name="code">Mã đề kiểm tra.</param>
        Task AppendGuideAsync(
            WordprocessingDocument doc,
            List<QuestionExport> answers,
            MixInfo mixInfo,
            string code
        );

        /// <summary>
        /// Di chuyển bảng bài luận (Essay Table) xuống cuối tài liệu Word.
        /// </summary>
        /// <param name="answerDoc">Tài liệu Word chứa bảng bài luận.</param>
        Task MoveEssayTableToEndAsync(WordprocessingDocument answerDoc);

        /// <summary>
        /// Thêm footer vào tài liệu Word theo phiên bản đề.
        /// </summary>
        /// <param name="doc">Tài liệu Word cần thêm footer.</param>
        /// <param name="version">Mã phiên bản đề kiểm tra.</param>
        Task AddFooterAsync(WordprocessingDocument doc, string version);
    }
}
