using ezMix.Helpers;
using ezMix.Services.Interfaces;
using Microsoft.Office.Interop.Word;
using MTGetEquationAddin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Word.Application;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace ezMix.Services.Implementations
{
    public class InteropWordService : IInteropWordService
    {
        private Application _wordApp;

        public async Task<Document> OpenDocumentAsync(string filePath, bool visible)
        {
            var wordApp = await GetWordAppAsync();
            wordApp.Visible = visible;

            return await Task.Run(() => wordApp.Documents.Open(filePath));
        }

        public async Task SaveDocumentAsync(_Document document)
        {
            try
            {
                await Task.Run(() => document.Save());
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi xảy ra ở hàm SaveDocumentAsync: {ex.Message}");
            }
        }

        // Đóng một document
        public async Task CloseDocumentAsync(Word._Document doc)
        {
            await Task.Run(() =>
            {
                try
                {
                    doc.Close(false); // false = không lưu
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm CloseDocumentAsync: {ex.Message}");
                }
                finally
                {
                    // Giải phóng COM object để tránh lỗi RPC
                    if (doc != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                    }
                }
            });
        }

        // Đóng tất cả document đang mở
        public async Task CloseAllDocumentsAsync()
        {
            if (_wordApp == null) return;

            await Task.Run(async () =>
            {
                try
                {
                    for (int i = _wordApp.Documents.Count; i >= 1; i--)
                    {
                        var doc = _wordApp.Documents[i];
                        await CloseDocumentAsync(doc);
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm CloseAllDocumentsAsync: {ex.Message}");
                }
            });
        }

        // Thoát Word Application
        public async Task QuitWordAppAsync()
        {
            if (_wordApp == null) return;

            await Task.Run(() =>
            {
                try
                {
                    _wordApp.Quit();
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm QuitWordAppAsync: {ex.Message}");
                }
                finally
                {
                    ReleaseComObject(_wordApp);
                    _wordApp = null;
                }
            });
        }

        public async Task FormatDocumentAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    var normalStyle = document.Styles["Normal"];
                    FormatNormalStyle(normalStyle);

                    var setup = document.PageSetup;
                    FormatPageSetup(setup);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm FormatDocumentAsync: {ex.Message}");
                }
            });
        }

        public async Task ReplaceAsync(
            _Document document,
            Dictionary<string, string> replacements,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                try
                {
                    var find = document.Content.Find;
                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = Type.Missing;

                    foreach (var kvp in replacements)
                    {
                        find.ClearFormatting();
                        find.Text = kvp.Key;
                        find.Replacement.ClearFormatting();
                        find.Replacement.Text = kvp.Value;

                        find.MatchCase = matchCase;
                        find.MatchWholeWord = matchWholeWord;
                        find.MatchWildcards = matchWildcards;

                        find.Execute(
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ReplaceAsync: {ex.Message}");
                }
            });
        }

        public async Task ReplaceFirstAsync(
            Paragraph paragraph,
            string findText,
            string replaceWithText,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                try
                {
                    var find = paragraph.Range.Find;
                    find.ClearFormatting();
                    find.Text = findText;
                    find.Replacement.ClearFormatting();
                    find.Replacement.Text = replaceWithText;

                    find.MatchCase = matchCase;
                    find.MatchWholeWord = matchWholeWord;
                    find.MatchWildcards = matchWildcards;

                    object replaceOne = Word.WdReplace.wdReplaceOne;
                    object missing = Type.Missing;

                    find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceOne, ref missing, ref missing, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ReplaceFirstAsync: {ex.Message}");
                }
            });
        }

        public async Task ReplaceUntilDoneAsync(
            _Document document,
            Dictionary<string, string> replacements,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false,
            int maxIterations = 100)
        {
            await Task.Run(() =>
            {
                try
                {
                    object missing = Type.Missing;
                    object replaceAll = Word.WdReplace.wdReplaceAll;

                    foreach (var pair in replacements)
                    {
                        string findText = pair.Key;
                        string replaceWithText = pair.Value;

                        int count = 0;
                        string previousText = document.Content.Text;

                        while (count < maxIterations)
                        {
                            var find = document.Content.Find;
                            find.ClearFormatting();
                            find.Text = findText;
                            find.Replacement.ClearFormatting();
                            find.Replacement.Text = replaceWithText;

                            find.MatchCase = matchCase;
                            find.MatchWholeWord = matchWholeWord;
                            find.MatchWildcards = matchWildcards;

                            find.Execute(
                                ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                            string currentText = document.Content.Text;
                            if (currentText == previousText)
                                break;

                            previousText = currentText;
                            count++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ReplaceUntilDoneAsync: {ex.Message}");
                }
            });
        }

        public async Task ReplaceInSectionAsync(
            _Document document,
            int sectionIndex,
            string findText,
            string replaceWithText,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                try
                {
                    if (sectionIndex < 1 || sectionIndex > document.Sections.Count)
                    {
                        MessageHelper.Error("Chỉ mục section không hợp lệ.");
                        return;
                    }

                    var find = document.Sections[sectionIndex].Range.Find;
                    find.ClearFormatting();
                    find.Text = findText;
                    find.Replacement.ClearFormatting();
                    find.Replacement.Text = replaceWithText;

                    find.MatchCase = matchCase;
                    find.MatchWholeWord = matchWholeWord;
                    find.MatchWildcards = matchWildcards;

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = Type.Missing;

                    find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ReplaceInSectionAsync: {ex.Message}");
                }
            });
        }

        public async Task ReplaceRedTextWithUnderlineAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    var find = document.Content.Find;

                    find.ClearFormatting();
                    find.Font.Color = Word.WdColor.wdColorRed;
                    find.Text = "";

                    find.Replacement.ClearFormatting();
                    find.Replacement.Font.Color = Word.WdColor.wdColorBlack;
                    find.Replacement.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    find.Replacement.Text = "^&"; // Giữ nguyên nội dung gốc

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = Type.Missing;
                    object format = true;

                    find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref format, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ReplaceRedTextWithUnderlineAsync: {ex.Message}");
                }
            });
        }

        public async Task ReplaceUnderlineWithRedTextAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    var find = document.Content.Find;

                    find.ClearFormatting();
                    find.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    find.Text = "";

                    find.Replacement.ClearFormatting();
                    find.Replacement.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    find.Replacement.Font.Color = Word.WdColor.wdColorRed;
                    find.Replacement.Text = "^&"; // Giữ nguyên nội dung gốc

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = Type.Missing;
                    object format = true;

                    find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref format, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ReplaceUnderlineWithRedTextAsync: {ex.Message}");
                }
            });
        }

        public async Task ReplaceInRangeAsync(
            _Document document,
            int start,
            int end,
            string findText,
            string replaceWithText,
            bool matchCase = false,
            bool matchWholeWord = false,
            bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                try
                {
                    if (start < 0 || end > document.Content.End || start > end)
                    {
                        MessageHelper.Error("Giá trị start hoặc end không hợp lệ.");
                        return;
                    }

                    var range = document.Range(start, end);
                    var find = range.Find;

                    find.ClearFormatting();
                    find.Text = findText;
                    find.Replacement.ClearFormatting();
                    find.Replacement.Text = replaceWithText;

                    find.MatchCase = matchCase;
                    find.MatchWholeWord = matchWholeWord;
                    find.MatchWildcards = matchWildcards;

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = Type.Missing;

                    find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ReplaceInRangeAsync: {ex.Message}");
                }
            });
        }

        public async Task ConvertListFormatToTextAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    var range = document.Content;
                    if (range.ListFormat.ListType != Word.WdListType.wdListNoNumbering)
                    {
                        range.ListFormat.ConvertNumbersToText();
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ConvertListFormatToTextAsync: {ex.Message}");
                }
            });
        }

        public async Task DeleteAllHeadersAndFootersAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    foreach (Word.Section section in document.Sections)
                    {
                        foreach (Word.HeaderFooter header in section.Headers)
                        {
                            if (header.Exists)
                                header.Range.Delete();
                        }

                        foreach (Word.HeaderFooter footer in section.Footers)
                        {
                            if (footer.Exists)
                                footer.Range.Delete();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm DeleteAllHeadersAndFootersAsync: {ex.Message}");
                }
            });
        }

        public async Task SetAnswersToABCDAsync(_Document document)
        {
            try
            {
                int questionIndex = 0, answerIndex = 0;

                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    string text = paragraph.Range.Text.Trim();

                    if (Constants.QuestionHeaderRegex.IsMatch(text))
                    {
                        questionIndex++;
                        answerIndex = 0;
                    }

                    if (text.Contains(Constants.ANSWER_TEMPLATE) || Constants.MultipleChoiceAnswerRegex.IsMatch(text))
                    {
                        string label = GenerateLabel(answerIndex); // A, B, C, D...
                        await ReplaceFirstAsync(paragraph, Constants.ANSWER_TEMPLATE, $"{label}. ");
                        answerIndex++;
                    }

                    // Nếu đáp án đang là a) → thay bằng nhãn đúng theo thứ tự
                    if (text.StartsWith("a)"))
                    {
                        string label = GenerateLabel(answerIndex).ToLower(); // a), b), c), d)
                        await ReplaceFirstAsync(paragraph, "a)", $"{label}) ");
                        answerIndex++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi xảy ra ở hàm SetAnswersToABCDAsync: {ex.Message}");
            }
        }

        public async Task SetQuestionsToNumberAsync(_Document document)
        {
            try
            {
                int questionIndex = 0;
                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    string text = paragraph.Range.Text.Trim();

                    if (await IsQuestionAsync(text))
                    {
                        questionIndex++;
                        await ReplaceFirstAsync(paragraph, Constants.QUESTION_TEMPLATE, $"Câu {questionIndex}: ");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi xảy ra ở hàm SetQuestionsToNumberAsync: {ex.Message}");
            }
        }

        public async Task FormatQuestionAndAnswerAsync(_Document document)
        {
            try
            {
                int soCauHoi = 0;
                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    string str = paragraph.Range.Text;
                    if (await IsQuestionAsync(str))
                    {
                        soCauHoi++;
                        if (str.StartsWith($"Câu {soCauHoi}:"))
                        {
                            int boldCount = $"Câu {soCauHoi}:".Length;
                            BoldCharacters(paragraph, boldCount);
                        }
                        else if (StartsWithAny(str, "<#>", "<G>", "<g>"))
                        {
                            BoldCharacters(paragraph, 3);
                        }
                        else if (StartsWithAny(str, "<NB>", "<TH>", "<VD>"))
                        {
                            BoldCharacters(paragraph, 4);
                        }
                        else if (str.StartsWith("<VDC>"))
                        {
                            BoldCharacters(paragraph, 5);
                        }
                        else if (str.StartsWith("#"))
                        {
                            BoldCharacters(paragraph, 1);
                        }
                        else if (str.StartsWith("[<br>]"))
                        {
                            BoldCharacters(paragraph, 6);
                        }
                    }
                    else if (StartsWithAny(str, "A.", "B.", "C.", "D.", "a)", "b)", "c)", "d)"))
                    {
                        FormatAnswer(paragraph, 2);
                    }
                    else if (str.StartsWith("<$>"))
                    {
                        FormatAnswer(paragraph, 3);
                    }
                    else
                    {
                        paragraph.Range.Font.Color = WdColor.wdColorBlack;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi xảy ra ở hàm FormatQuestionAndAnswerAsync: {ex.Message}");
            }
        }

        public async Task UpdateFieldsAsync(string filePath)
        {
            Word.Application app = null;
            Word._Document document = null;
            try
            {
                app = new Word.Application();
                document = app.Documents.Open(filePath, ReadOnly: false, Visible: false);
                document.Fields.Update();
                document.Save();
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi xảy ra ở hàm UpdateFieldsAsync: {ex.Message}");
            }
            finally
            {
                if (document != null)
                {
                    document.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(document);
                }
                if (app != null)
                {
                    app.Quit(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                }
            }
        }

        public async Task ClearTabStopsAsync(Word.Paragraph paragraph)
        {
            await Task.Run(() =>
            {
                try
                {
                    var tabStops = paragraph.Format.TabStops;
                    for (int i = tabStops.Count; i >= 1; i--)
                    {
                        tabStops[i].Clear();
                    }
                    ReleaseComObject(tabStops);
                    tabStops = null;
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ClearTabStopsAsync: {ex.Message}");
                }
            });
        }

        public async Task<int> FixMathTypeAsync(_Document document)
        {
            return await Task.Run(() =>
            {
                try
                {
                    Word.InlineShapes shapes = document?.InlineShapes;
                    if (shapes != null && shapes.Count > 0)
                    {
                        var connect = new Connect();
                        int numShapesIterated = connect.IterateShapes(ref shapes, true, true);
                        return numShapesIterated;
                    }
                    return 0;
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm FixMathTypeAsync: {ex.Message}");
                    return -1;
                }
            });
        }

        public async Task<int> ConvertEquationToMathTypeAsync(_Document document)
        {
            return await Task.Run(() =>
            {
                int convertedCount = 0;

                try
                {
                    var omaths = document?.OMaths;
                    if (omaths == null || omaths.Count == 0)
                        return 0;

                    var connect = new Connect();

                    foreach (Word.OMath omath in omaths)
                    {
                        Word.Range range = omath.Range;

                        // Chèn đối tượng MathType tại vị trí công thức Equation
                        Word.InlineShape shape = document.InlineShapes.AddOLEObject(
                            ClassType: "Equation.DSMT4",
                            FileName: "",
                            LinkToFile: false,
                            DisplayAsIcon: false,
                            Range: range);

                        // Lấy GUID và verb để gửi dữ liệu
                        Guid clsid;
                        string progID = "Equation.DSMT4";
                        if (!connect.FindAutoConvert(ref progID, out clsid))
                            continue;

                        if (!connect.DoesServerExist(ref clsid))
                            continue;

                        string format = "MathML";
                        if (!connect.DoesServerSupportFormat(ref clsid, ref format))
                            continue;

                        int verbIndex = connect.GetVerbIndex("RunForConversion", ref clsid);
                        if (verbIndex == 999)
                            continue;

                        // Gửi MathML mẫu vào đối tượng MathType
                        connect.Equation_SetData(ref shape, ref format, verbIndex, true);
                        convertedCount++;
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ConvertEquationToMathTypeAsync: {ex.Message}");
                    return -1;
                }

                return convertedCount;
            });
        }

        public async Task RejectAllChangesAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    // Kiểm tra nếu có thay đổi đang được theo dõi
                    if (document.Revisions.Count > 0)
                    {
                        document.Revisions.RejectAll();
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm RejectAllChangesAsync: {ex.Message}");
                }
                finally
                {
                    Marshal.FinalReleaseComObject(document.Revisions);
                }
            });
        }

        public void NormalizeParagraphEnds(_Document document)
        {
            foreach (Word.Paragraph para in document.Paragraphs)
            {
                Word.Range range = para.Range;
                if (range.Characters.Count > 0)
                {
                    Word.Range lastChar = range.Characters.Last;
                    string text = lastChar.Text;

                    // Nếu ký tự cuối là Enter
                    if (text == "\r" || text == "\n")
                    {
                        if (lastChar.Font.Subscript == -1)
                        {
                            lastChar.Font.Subscript = 0; // Bỏ subscript
                        }
                        if (lastChar.Font.Superscript == -1)
                        {
                            lastChar.Font.Superscript = 0; // Bỏ superscript
                        }
                    }
                }
            }
        }

        private async Task<Application> GetWordAppAsync()
        {
            if (_wordApp == null)
            {
                _wordApp = await Task.Run(() => new Application());
            }
            return _wordApp;
        }
        private static float CmToPt(float cm) => (float)(cm * 28.35);
        private static void FormatNormalStyle(Style style)
        {
            var para = style.ParagraphFormat;
            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            para.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
            para.LeftIndent = 0f;
            para.CharacterUnitLeftIndent = 0f;
            para.RightIndent = 0f;
            para.SpaceBefore = 0f;
            para.SpaceAfter = 0f;
            para.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
            para.LineSpacing = 14.4f;
            para.FirstLineIndent = 0f;
            para.HangingPunctuation = 0;
        }
        private static void FormatPageSetup(PageSetup setup)
        {
            setup.PaperSize = Word.WdPaperSize.wdPaperA4;
            setup.Orientation = Word.WdOrientation.wdOrientPortrait;
            setup.TopMargin = CmToPt(1.27f);
            setup.BottomMargin = CmToPt(1.27f);
            setup.LeftMargin = CmToPt(1.27f);
            setup.RightMargin = CmToPt(1.27f);
            setup.HeaderDistance = CmToPt(1.27f);
            setup.FooterDistance = CmToPt(1.27f);
        }
        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                try
                {
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        Marshal.FinalReleaseComObject(comObject);
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi xảy ra ở hàm ReleaseComObject: {ex.Message}");
                }
            }
        }
        private static Task<bool> IsQuestionAsync(string s)
        {
            bool result = Constants.QuestionPrefixes.Any(s.StartsWith) || Constants.QuestionHeaderRegex.IsMatch(s);
            return Task.FromResult(result);
        }
        private static Task<bool> IsAnswerAsync(string s)
        {
            bool result = Constants.AnswerPrefixes.Any(s.StartsWith);
            return Task.FromResult(result);
        }
        private static string GenerateLabel(int index)
        {
            index++; // Bắt đầu từ 1 thay vì 0
            var label = new StringBuilder();

            while (index > 0)
            {
                index--;
                label.Insert(0, (char)('A' + index % 26));
                index /= 26;
            }

            return label.ToString();
        }
        private void BoldCharacters(Word.Paragraph paragraph, int count)
        {
            for (int i = 1; i <= count; i++)
            {
                var font = paragraph.Range.Characters[i].Font;
                font.Bold = 1;                                   // in đậm
                font.Italic = 0;                                 // không in nghiêng
                font.Underline = Word.WdUnderline.wdUnderlineNone; // không gạch chân
            }
        }

        private void FormatAnswer(Word.Paragraph paragraph, int boldCount)
        {
            var range = paragraph.Range;
            int totalChars = range.Characters.Count;

            // Đặt toàn bộ đoạn về màu đen, bỏ đậm
            range.Font.Color = WdColor.wdColorBlack;
            range.Font.Bold = 0;

            // Kiểm tra đủ độ dài
            if (totalChars >= boldCount)
            {
                bool isUnderlined = range.Characters[1].Font.Underline == WdUnderline.wdUnderlineSingle;
                for (int i = 1; i <= boldCount; i++)
                {
                    range.Characters[i].Font.Bold = 1;
                    if (isUnderlined)
                    {
                        range.Characters[i].Font.Underline = WdUnderline.wdUnderlineSingle;
                    }
                }

                // Bỏ gạch chân phần sau nếu không cần giữ
                Word.Range tailRange = range.Duplicate;
                tailRange.MoveStart(WdUnits.wdCharacter, boldCount);
                tailRange.Font.Underline = WdUnderline.wdUnderlineNone;
            }
        }
        private bool StartsWithAny(string input, params string[] prefixes)
        {
            return prefixes.Any(p => input.StartsWith(p));
        }
    }
}
