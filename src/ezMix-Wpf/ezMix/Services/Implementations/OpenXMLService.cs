using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using ezMix.Helpers;
using ezMix.Models;
using ezMix.Models.Enums;
using ezMix.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Group = DocumentFormat.OpenXml.Vml.Group;
using InsideHorizontalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder;
using InsideVerticalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using Level = ezMix.Models.Enums.Level;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using SectionProperties = DocumentFormat.OpenXml.Wordprocessing.SectionProperties;
using Shape = DocumentFormat.OpenXml.Vml.Shape;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TabStop = DocumentFormat.OpenXml.Wordprocessing.TabStop;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;


namespace ezMix.Services.Implementations
{
    public class OpenXMLService : IOpenXMLService
    {
        private readonly IInteropWordService _interopWordService;
        private readonly IExcelService _excelService;

        public OpenXMLService(IInteropWordService interopWordService, IExcelService excelAnswerExporter)
        {
            _interopWordService = interopWordService;
            _excelService = excelAnswerExporter;
        }

        public Task<WordprocessingDocument> CreateDocumentAsync(string filePath)
        {
            var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            document.AddMainDocumentPart();
            document.MainDocumentPart.Document = new Document(new Body());
            document.MainDocumentPart.Document.Save();
            return Task.FromResult(document);
        }

        public Task<WordprocessingDocument> OpenDocumentAsync(string filePath, bool isEditTable)
        {
            return Task.FromResult(WordprocessingDocument.Open(filePath, isEditTable));
        }

        public Task SaveDocumentAsync(WordprocessingDocument document)
        {
            document.MainDocumentPart.Document.Save();
            return Task.CompletedTask;
        }

        public Task CloseDocumentAsync(WordprocessingDocument document)
        {
            document.Dispose();
            return Task.CompletedTask;
        }

        public Task<string> GetDocumentTextAsync(WordprocessingDocument document)
        {
            return Task.FromResult(document.MainDocumentPart.Document.Body.InnerText);
        }

        public Task<Body> GetDocumentBodyAsync(WordprocessingDocument document)
        {
            return Task.FromResult(document.MainDocumentPart.Document.Body);
        }

        public Task FormatAllParagraphsAsync(WordprocessingDocument doc, MixInfo mixInfo)
        {
            var body = doc.MainDocumentPart.Document.Body;
            foreach (var para in body.Descendants<Paragraph>())
                FormatParagraph(para, mixInfo);
            return Task.CompletedTask;
        }

        public async Task<List<QuestionExport>> ShuffleQuestionsAsync(WordprocessingDocument doc, string version, WordprocessingDocument answerDoc = null, MixInfo mixInfo = null)
        {
            return await Task.Run(async () =>
            {
                var body = doc.MainDocumentPart.Document.Body;

                var allBlocks = SplitQuestions(body);
                var grouped = new Dictionary<QuestionType, List<List<OpenXmlElement>>>
                {
                    { QuestionType.MultipleChoice, new List<List<OpenXmlElement>>() },
                    { QuestionType.TrueFalse, new List<List<OpenXmlElement>>() },
                    { QuestionType.ShortAnswer, new List<List<OpenXmlElement>>() },
                    { QuestionType.Essay, new List<List<OpenXmlElement>>() }
                };

                foreach (var block in allBlocks)
                {
                    var type = DetectQuestionType(block);
                    grouped[type].Add(block);
                }

                var rng = new Random();
                foreach (var key in grouped.Keys.ToList())
                {
                    if (version.Equals(Constants.ROOT_CODE)) continue;

                    if (key == QuestionType.MultipleChoice && mixInfo?.IsShuffledQuestionMultipleChoice == false)
                        continue;

                    if (key == QuestionType.TrueFalse && mixInfo?.IsShuffledQuestionTrueFalse == false)
                        continue;

                    if (key == QuestionType.ShortAnswer && mixInfo?.IsShuffledShortAnswer == false)
                        continue;

                    if (key == QuestionType.Essay && mixInfo?.IsShuffledEssay == false)
                        continue;

                    grouped[key] = grouped[key].OrderBy(_ => rng.Next()).ToList();
                }

                var answers = new List<QuestionExport>();
                body.RemoveAllChildren();
                int questionNumber = 0;
                int index = 0;
                // ➤ Thêm biến này để theo dõi phân phối đáp án đúng
                int[] correctDistribution = new int[4]; // A,B,C,D

                string[] points = new string[]
                {
                    mixInfo.PointMultipleChoice,
                    mixInfo.PointTrueFalse,
                    mixInfo.PointShortAnswer,
                    mixInfo.PointEssay
                };

                foreach (var group in grouped.OrderBy(g => g.Key).Where(g => g.Value.Any()))
                {
                    int localQuestion = 0;
                    index++;

                    string title = CreateSectionTitle(group.Key, index, group.Value.Count, points);
                    if (!string.IsNullOrEmpty(title))
                    {
                        var heading = new Paragraph();
                        var parts = title.Split(new[] { '.' }, 3);
                        if (parts.Length >= 3)
                        {
                            var boldRun = new Run(new RunProperties(new Bold()), new Text($"{parts[0]}.{parts[1]}.") { Space = SpaceProcessingModeValues.Preserve });
                            var normalRun = new Run(new Text(parts[2]) { Space = SpaceProcessingModeValues.Preserve });
                            heading.Append(boldRun, normalRun);
                        }
                        body.Append(heading);
                    }

                    foreach (var block in group.Value)
                    {
                        localQuestion++;
                        questionNumber++;
                        var type = group.Key;
                        var newBlock = ShuffleAnswers(block, type, version, doc.MainDocumentPart, out string correct, out var answerElements, mixInfo, correctDistribution);

                        // Cập nhật số thứ tự câu hỏi hiển thị
                        var firstPara = newBlock.OfType<Paragraph>().FirstOrDefault();
                        if (firstPara != null)
                        {
                            await UpdateQuestionNumberAsync(firstPara, localQuestion);

                            // Đảm bảo nhãn "Câu *:" in đậm
                            var labelText = firstPara.Descendants<Text>().FirstOrDefault(t =>
                                System.Text.RegularExpressions.Regex.IsMatch(t.Text.Trim(), @"^Câu\s+\d+:"));
                            if (labelText != null)
                            {
                                var labelRun = labelText.Parent as Run;
                                if (labelRun != null)
                                {
                                    if (labelRun.RunProperties == null)
                                        labelRun.RunProperties = new RunProperties();

                                    labelRun.RunProperties.Bold = new Bold() { Val = OnOffValue.FromBoolean(true) };
                                }
                            }
                        }


                        // Lấy điểm cho câu hỏi tự luận
                        string point = null;
                        if (type == QuestionType.Essay)
                        {
                            // Tìm điểm trong toàn bộ block (bao gồm cả đáp án)
                            var blockText = string.Join(" ", block.SelectMany(el => el.Descendants<Run>().SelectMany(run => run.Elements<Text>().Select(t => t.Text))));
                            point = ExtractPointFromText(blockText);

                            // ➤ Tách phần `answerElements` ra từ block
                            var allParas = block.OfType<Paragraph>().ToList();
                            var firstAnswerPara = allParas.FirstOrDefault(p => System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"));
                            if (firstAnswerPara != null)
                            {
                                answerElements = new List<OpenXmlElement> { firstAnswerPara };
                                int idx = block.IndexOf(firstAnswerPara);
                                for (int i = idx + 1; i < block.Count; i++)
                                {
                                    if (block[i] is Paragraph p && System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+")) break;
                                    answerElements.Add(block[i]);
                                }
                            }
                        }

                        foreach (var el in newBlock)
                            body.Append(el.CloneNode(true));

                        // ✅ Nếu có đáp án tự luận và đang có file đáp án mở
                        if (type == QuestionType.Essay && answerDoc != null)
                        {
                            var sourcePart = doc.MainDocumentPart;
                            var targetPart = answerDoc.MainDocumentPart;
                            var answerBody = targetPart.Document.Body ?? targetPart.Document.AppendChild(new Body());

                            // Tạo bảng nếu chưa có
                            var existingTable = answerBody.Elements<Table>().FirstOrDefault(t =>
                                t.Descendants<TableRow>().Any(r => r.InnerText.Contains("Câu") && r.InnerText.Contains("Đáp án")));
                            Table table = existingTable ?? CreateTable();
                            if (existingTable == null)
                            {
                                var header = new TableRow();
                                header.Append(CreateCell("Câu", "700"), CreateCell("Đáp án", "5000"), CreateCell("Điểm", "700"));
                                table.Append(header);
                                answerBody.Append(table);
                            }

                            // ➤ Tách phần đáp án tự luận
                            var extracted = await ExtractEssayAnswerAsync(block);

                            // ➤ Clone nguyên block đáp án (ảnh + công thức MathType + VML + strip nhãn A.)
                            var contentClones = CloneAnswerBlock(extracted, sourcePart, targetPart);

                            // Trích điểm nếu có
                            var fullText = string.Join(" ", block.Select(b => b.InnerText));
                            point = ExtractPointFromText(fullText);

                            // Thêm hàng mới vào bảng đáp án
                            var row = new TableRow();
                            row.Append(CreateCell(localQuestion.ToString(), "700"),
                                       CreateCell(contentClones, "5000"),
                                       CreateCell(point ?? string.Empty, "700"));
                            table.Append(row);

                            // Không dùng lại answerElements nữa
                            answerElements = null;
                        }

                        // Ghi vào danh sách câu hỏi
                        answers.Add(new QuestionExport
                        {
                            QuestionNumber = localQuestion,
                            CorrectAnswer = correct,
                            Type = type,
                            Point = point,
                            AnswerElements = answerElements
                        });
                    }
                }
                doc.MainDocumentPart.Document.Save();
                return answers;
            });
        }

        private List<OpenXmlElement> ShuffleAnswers(
            List<OpenXmlElement> block,
            QuestionType type,
            string version,
            MainDocumentPart sourcePart,
            out string correctAnswer,
            out List<OpenXmlElement> answerElements,
            MixInfo mixInfo, int[] correctDistribution)
        {
            switch (type)
            {
                case QuestionType.MultipleChoice:
                    return ShuffleMultipleChoice(block, version, sourcePart, out correctAnswer, out answerElements, mixInfo, correctDistribution);

                case QuestionType.TrueFalse:
                    return ShuffleTrueFalse(block, version, sourcePart, out correctAnswer, out answerElements, mixInfo);

                case QuestionType.ShortAnswer:
                    return ShuffleShortAnswer(block, version, sourcePart, out correctAnswer, out answerElements, mixInfo);

                case QuestionType.Essay:
                    return ShuffleEssay(block, version, sourcePart, out correctAnswer, out answerElements, mixInfo);

                default:
                    correctAnswer = string.Empty;
                    answerElements = null;
                    return block;
            }
        }

        private int FindCorrectIndex(List<List<OpenXmlElement>> shuffled)
        {
            for (int i = 0; i < shuffled.Count; i++)
            {
                var para = shuffled[i].OfType<Paragraph>().FirstOrDefault();
                if (para != null)
                {
                    // Nếu trong nhóm có Run được gạch chân (underline) => đó là đáp án đúng
                    bool isCorrect = para.Descendants<Run>()
                        .Any(r => r.RunProperties?.Underline?.Val != null &&
                                  r.RunProperties.Underline.Val != UnderlineValues.None);
                    if (isCorrect) return i;
                }
            }
            return -1;
        }

        private List<OpenXmlElement> ShuffleMultipleChoice(
     List<OpenXmlElement> block,
     string version,
     MainDocumentPart sourcePart,
     out string correctAnswer,
     out List<OpenXmlElement> answerElements,
     MixInfo mixInfo, int[] correctDistribution)
        {
            var rnd = new Random();
            correctAnswer = string.Empty;
            answerElements = null;

            var allElements = block;
            var allParas = allElements.OfType<Paragraph>().ToList();

            var answerStartParas = allParas
                .Where(p => System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[A-D]\."))
                .ToList();

            if (answerStartParas.Count < 2)
                return allElements.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();

            // Gom nhóm đáp án
            var answerGroups = new List<List<OpenXmlElement>>();
            for (int i = 0; i < answerStartParas.Count; i++)
            {
                var startPara = answerStartParas[i];
                int startIndex = allElements.IndexOf(startPara);
                int endIndex = (i < answerStartParas.Count - 1)
                    ? allElements.IndexOf(answerStartParas[i + 1])
                    : allElements.Count;
                answerGroups.Add(allElements.Skip(startIndex).Take(endIndex - startIndex).ToList());
            }

            // Phần câu hỏi
            int firstAnswerIndex = allElements.IndexOf(answerStartParas.First());
            var questionElements = allElements.Take(firstAnswerIndex).ToList();

            // Shuffle nếu cần
            var shuffled = (version.Equals(Constants.ROOT_CODE) || mixInfo?.IsShuffledAnswerMultipleChoice == false)
                ? answerGroups
                : answerGroups.OrderBy(_ => rnd.Next()).ToList();

            if (!version.Equals(Constants.ROOT_CODE))
            {
                int correctIndex = FindCorrectIndex(shuffled);
                if (correctIndex >= 0)
                {
                    int max = correctDistribution.Max();
                    int min = correctDistribution.Min();

                    if (correctDistribution[correctIndex] > min + 2)
                    {
                        var targetIndex = correctDistribution
                            .Select((count, idx) => new { count, idx })
                            .OrderBy(x => x.count)
                            .First().idx;

                        var correctGroup = shuffled[correctIndex];
                        shuffled.RemoveAt(correctIndex);
                        shuffled.Insert(targetIndex, correctGroup);
                        correctIndex = targetIndex;
                    }

                    correctDistribution[correctIndex]++;
                }
            }

            var labels = new[] { "A.", "B.", "C.", "D." };
            correctAnswer = UpdateAnswerLabels(shuffled, labels);

            bool hasMultiElementAnswer = shuffled.Any(g =>
                g.Where(e =>
                    !(e is BookmarkStart) &&
                    !(e is BookmarkEnd) &&
                    !(e is CommentRangeStart) &&
                    !(e is CommentRangeEnd) &&
                    !(e is CommentReference) &&
                    !(e is ProofError) &&
                    !(e is PermStart) &&
                    !(e is PermEnd)
                ).Count() > 1);

            bool hasMathType = shuffled.Any(g => g.OfType<Paragraph>().Any(p => p.Descendants<OleObject>().Any() || p.Descendants<EmbeddedObject>().Any()));
            bool hasEquation = shuffled.Any(g => g.OfType<Paragraph>().Any(p => p.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any()));
            bool hasImage = shuffled.Any(g => g.OfType<Paragraph>().Any(p => p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any()));

            int maxLength = shuffled.Select(g =>
            {
                var para = g.OfType<Paragraph>().FirstOrDefault();
                if (para == null) return 0;

                int textLen = (para.InnerText ?? string.Empty).Trim().Length;
                int eqLen = para.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>()
                                .SelectMany(m => m.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                                .Sum(t => (t.Text ?? string.Empty).Length);

                return textLen + eqLen;
            }).DefaultIfEmpty().Max();

            long maxImageWidth = 0L;
            foreach (var g in shuffled)
            {
                var para = g.OfType<Paragraph>().FirstOrDefault();
                if (para == null) continue;

                var wpDrawing = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();
                var inline = wpDrawing?.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().FirstOrDefault();
                var anchor = wpDrawing?.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();

                var extent = inline?.Extent ?? anchor?.Extent;
                if (extent != null && extent.Cx.HasValue)
                {
                    maxImageWidth = Math.Max(maxImageWidth, extent.Cx.Value);
                }
            }

            int perLine;
            if (hasMathType)
            {
                perLine = 2; // MathType hoặc MathType + text
            }
            else if (hasMultiElementAnswer)
            {
                perLine = 1;
            }
            else if (hasEquation)
            {
                if (maxLength < 15) perLine = 4;
                else if (maxLength < 30) perLine = 2;
                else perLine = 1;
            }
            else if (hasImage)
            {
                // Dựa vào độ rộng ảnh
                if (maxImageWidth < 2000000) perLine = 4;   // ảnh nhỏ
                else if (maxImageWidth < 4000000) perLine = 2; // ảnh vừa
                else perLine = 1; // ảnh lớn
            }
            else
            {
                perLine = maxLength < 25 ? 4 :
                          maxLength < 50 ? 2 : 1;
            }


            int[] tabPositions;
            if (perLine == 1)
                tabPositions = new int[] { Constants.TABSTOP_1 };
            else if (perLine == 2)
                tabPositions = new int[] { Constants.TABSTOP_1, Constants.TABSTOP_3 };
            else if (perLine == 3)
                tabPositions = new int[] { Constants.TABSTOP_1, Constants.TABSTOP_2, Constants.TABSTOP_3 };
            else
                tabPositions = new int[] { Constants.TABSTOP_1, Constants.TABSTOP_2, Constants.TABSTOP_3, Constants.TABSTOP_4 };

            var result = new List<OpenXmlElement>();
            result.AddRange(questionElements.Select(e => e.CloneNode(true)));

            // Hiển thị đáp án theo perLine
            for (int i = 0; i < shuffled.Count; i += perLine)
            {
                var lineGroups = shuffled.Skip(i).Take(perLine).ToList();

                var para = new Paragraph
                {
                    ParagraphProperties = new ParagraphProperties(new Tabs(
                        tabPositions.Select(tp => new TabStop() { Val = TabStopValues.Left, Position = tp })))
                };

                for (int j = 0; j < lineGroups.Count; j++)
                {
                    para.Append(new Run(new TabChar()));

                    var clonedGroup = lineGroups[j].Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();
                    var firstPara = clonedGroup.OfType<Paragraph>().FirstOrDefault();

                    if (firstPara != null)
                    {
                        // Tạo nhãn mới
                        var labelRun = new Run(new RunProperties(new Bold()),
                            new Text(labels[i + j] + " ") { Space = SpaceProcessingModeValues.Preserve });
                        para.Append(labelRun);

                        // Append tất cả child node của Paragraph
                        foreach (var node in firstPara.Elements())
                        {
                            var clonedNode = node.CloneNode(true);

                            // Nếu là Run chứa nhãn cũ thì strip nhãn
                            if (clonedNode is Run run)
                            {
                                foreach (var textNode in run.Elements<Text>())
                                {
                                    textNode.Text = System.Text.RegularExpressions.Regex.Replace(textNode.Text, @"^[A-D]\.", "").TrimStart();
                                }
                                if (run.RunProperties?.Underline != null)
                                    run.RunProperties.Underline.Val = UnderlineValues.None;
                                para.Append(run);
                            }
                            // 👉 Nếu là OfficeMath thì giữ nguyên
                            else if (clonedNode is DocumentFormat.OpenXml.Math.OfficeMath)
                            {
                                para.Append(clonedNode);
                            }
                            else
                            {
                                para.Append(clonedNode);
                            }
                        }
                    }

                    // Các element bổ sung trong group thì append vào cùng Paragraph
                    foreach (var el in clonedGroup.Skip(1))
                    {
                        para.Append(el.CloneNode(true));
                    }
                }

                result.Add(para);
            }

            return result;
        }

        private List<OpenXmlElement> ShuffleTrueFalse(
     List<OpenXmlElement> block,
     string version,
     MainDocumentPart sourcePart,
     out string correctAnswer,
     out List<OpenXmlElement> answerElements,
     MixInfo mixInfo)
        {
            var rnd = new Random();
            correctAnswer = string.Empty;
            answerElements = null;

            var allElements = block;
            var allParas = allElements.OfType<Paragraph>().ToList();

            // Tìm các Paragraph bắt đầu bằng a), b), c), d)
            var trueFalseFirstParas = allParas
                .Where(p => System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[a-d]\)"))
                .ToList();

            if (!trueFalseFirstParas.Any())
                return allElements.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();

            // Gom nhóm đáp án
            var answerGroups = new List<List<OpenXmlElement>>();
            for (int i = 0; i < trueFalseFirstParas.Count; i++)
            {
                var startPara = trueFalseFirstParas[i];
                int startIndex = allElements.IndexOf(startPara);
                int endIndex = (i < trueFalseFirstParas.Count - 1)
                    ? allElements.IndexOf(trueFalseFirstParas[i + 1])
                    : allElements.Count;
                answerGroups.Add(allElements.Skip(startIndex).Take(endIndex - startIndex).ToList());
            }

            // Metadata đáp án
            var meta = answerGroups.Select(g => new
            {
                Group = g,
                IsCorrect = g.Any(el => el.Descendants<Run>().Any(r =>
                    r.RunProperties?.Underline?.Val != null &&
                    r.RunProperties.Underline.Val != UnderlineValues.None))
            }).ToList();

            // Shuffle nếu cần
            var shuffled = (version.Equals(Constants.ROOT_CODE) || mixInfo?.IsShuffledAnswerTrueFalse == false)
                ? meta
                : meta.OrderBy(_ => rnd.Next()).ToList();

            var labels = new[] { "a)", "b)", "c)", "d)" };

            answerElements = new List<OpenXmlElement>();
            var correctTokens = new List<string>();
            var resultParas = new List<OpenXmlElement>();

            for (int i = 0; i < shuffled.Count && i < labels.Length; i++)
            {
                var clonedGroup = shuffled[i].Group.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();
                var firstPara = clonedGroup.OfType<Paragraph>().FirstOrDefault();

                // Tạo Paragraph mới với TabStop_1
                var para = new Paragraph
                {
                    ParagraphProperties = new ParagraphProperties(
                        new Tabs(new TabStop { Val = TabStopValues.Left, Position = Constants.TABSTOP_1 }))
                };

                // TabChar đầu dòng
                para.Append(new Run(new TabChar()));

                // Nhãn a) in đậm
                var labelRun = new Run(new Text(labels[i] + " "));
                labelRun.RunProperties = new RunProperties(new Bold());
                para.Append(labelRun);

                // Copy nội dung gốc nhưng bỏ nhãn cũ
                if (firstPara != null)
                {
                    // Xóa nhãn cũ trong firstPara
                    var oldLabel = firstPara.Descendants<Text>().FirstOrDefault(t =>
                        System.Text.RegularExpressions.Regex.IsMatch(t.Text.Trim(), @"^[a-d]\)"));
                    oldLabel?.Parent?.Remove();

                    // Append phần còn lại (giữ nguyên bảng, hình, công thức…)
                    foreach (var node in firstPara.Elements())
                        para.Append(node.CloneNode(true));
                }

                // Thêm vào kết quả
                resultParas.Add(para);

                // Các element bổ sung trong group (nếu có nhiều hơn 1 Paragraph)
                foreach (var el in clonedGroup.Skip(1))
                    resultParas.Add(el);


                // Cho file đáp án
                answerElements.AddRange(clonedGroup);
                correctTokens.Add($"{labels[i]} {(shuffled[i].IsCorrect ? "Đúng" : "Sai")}");
            }

            correctAnswer = string.Join(" ", correctTokens);

            // Phần câu hỏi (trước block đáp án)
            var firstAnswerPara = trueFalseFirstParas.First();
            var questionElements = allElements.Take(allElements.IndexOf(firstAnswerPara)).ToList();

            var result = new List<OpenXmlElement>();
            result.AddRange(questionElements.Select(e => e.CloneNode(true)));
            result.AddRange(resultParas);

            return result;
        }


        private List<OpenXmlElement> ShuffleShortAnswer(
     List<OpenXmlElement> block,
     string version,
     MainDocumentPart sourcePart,
     out string correctAnswer,
     out List<OpenXmlElement> answerElements,
     MixInfo mixInfo)
        {
            correctAnswer = string.Empty;
            answerElements = null;

            var allParas = block.OfType<Paragraph>().ToList();

            // Tìm Paragraph chứa đáp án ngắn (bắt đầu bằng A. ...)
            var para = allParas.FirstOrDefault(p =>
                System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"));

            if (para != null)
            {
                // Lấy toàn bộ text trong Paragraph
                var fullText = para.InnerText.Trim();

                // Loại bỏ nhãn "A. " ở đầu
                correctAnswer = System.Text.RegularExpressions.Regex.Replace(fullText, @"^[A-Z]\.\s*", "");

                // Trả lại tất cả element trừ phần chứa đáp án
                return block.Except(new[] { para })
                            .Select(e => (OpenXmlElement)e.CloneNode(true))
                            .ToList();
            }

            // Không có đáp án -> trả nguyên block
            return block.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();
        }


        private List<OpenXmlElement> ShuffleEssay(
    List<OpenXmlElement> block,
    string version,
    MainDocumentPart sourcePart,
    out string correctAnswer,
    out List<OpenXmlElement> answerElements,
    MixInfo mixInfo)
        {
            correctAnswer = string.Empty;
            answerElements = null;

            var allParas = block.OfType<Paragraph>().ToList();
            var firstAnswerPara = allParas.FirstOrDefault(p =>
                System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"));

            if (firstAnswerPara == null)
            {
                // Không có đáp án, trả nguyên block (câu hỏi)
                return block.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();
            }

            // Tách phần đáp án
            int idx = block.IndexOf(firstAnswerPara);
            var answerBlock = block.Skip(idx).ToList();
            var questionBlock = block.Take(idx).ToList();

            // Clone nguyên xi đáp án để đưa sang file đáp án
            answerElements = answerBlock.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();

            // Ghép nội dung text để lưu CorrectAnswer
            correctAnswer = string.Join(Environment.NewLine, answerElements.Select(e => e.InnerText.Trim()));

            // Trả lại phần câu hỏi (không gồm đáp án) để append vào đề trộn
            return questionBlock.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();
        }

        public async Task<List<Question>> ParseDocxQuestionsAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                var questions = new List<Question>();

                try
                {
                    using (var doc = WordprocessingDocument.Open(filePath, false))
                    {
                        var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

                        System.Text.RegularExpressions.Regex questionHeader = Constants.QuestionHeaderRegex;
                        System.Text.RegularExpressions.Regex mcAnswerRegex = Constants.MultipleChoiceAnswerRegex;
                        System.Text.RegularExpressions.Regex trueFalseAnswer = Constants.TrueFalseAnswerRegex;
                        System.Text.RegularExpressions.Regex levelRegex = Constants.LevelRegex;

                        int code = 1;
                        for (int i = 0; i < paragraphs.Count; i++)
                        {
                            string text = paragraphs[i].InnerText.Trim();
                            if (!questionHeader.IsMatch(text)) continue;

                            var question = new Question
                            {
                                Code = code++,
                                Level = Level.Know,
                                IsValid = false
                            };

                            var levelMatch = levelRegex.Match(text);
                            if (levelMatch.Success)
                            {
                                question.Level = GetLevelFromText(levelMatch.Value);
                                text = levelRegex.Replace(text, "").Trim();
                            }

                            var questionTextRuns = new List<Run>();
                            questionTextRuns.AddRange(paragraphs[i].Elements<Run>());
                            int j = i + 1;

                            while (j < paragraphs.Count && !questionHeader.IsMatch(paragraphs[j].InnerText.Trim()))
                            {
                                var para = paragraphs[j];
                                string line = para.InnerText.Trim();

                                if (mcAnswerRegex.IsMatch(line) || trueFalseAnswer.IsMatch(line))
                                    break;

                                questionTextRuns.AddRange(para.Elements<Run>());
                                j++;
                            }

                            string rawText = string.Join("", questionTextRuns.Select(r => r.InnerText).Where(t => !string.IsNullOrWhiteSpace(t))).Trim();
                            var match = System.Text.RegularExpressions.Regex.Match(rawText, @"^(Câu\s+\d+[\.:]?)\s*(.*)", RegexOptions.IgnoreCase);
                            question.QuestionText = match.Success ? $"{match.Groups[1].Value} {match.Groups[2].Value.Trim()}" : rawText;

                            var answers = new List<string>();
                            var correctAnswers = new List<string>();
                            var currentAnswerLines = new List<string>();

                            while (j < paragraphs.Count && !questionHeader.IsMatch(paragraphs[j].InnerText.Trim()))
                            {
                                var para = paragraphs[j];
                                string line = para.InnerText.Trim();

                                bool isNewAnswer = mcAnswerRegex.IsMatch(line) || trueFalseAnswer.IsMatch(line);

                                if (isNewAnswer)
                                {
                                    if (currentAnswerLines.Count > 0)
                                    {
                                        answers.Add(string.Join("\n", currentAnswerLines));
                                        currentAnswerLines.Clear();
                                    }

                                    currentAnswerLines.Add(line);

                                    string label = line.Substring(0, 2);
                                    if (trueFalseAnswer.IsMatch(line))
                                    {
                                        string indicator = IsUnderlined(para, label) ? "Đ" : "S";
                                        var levelMatchItem = levelRegex.Match(line);
                                        var level = levelMatchItem.Success ? GetLevelFromText(levelMatchItem.Value) : Level.Know;
                                        string formatted = $"{label} {indicator} ({ShortLevelCode(level)})";
                                        correctAnswers.Add(formatted);
                                    }
                                    else if (mcAnswerRegex.IsMatch(line) && IsUnderlined(para, label))
                                    {
                                        correctAnswers.Add(label.TrimEnd('.', ')'));
                                    }
                                }
                                else
                                {
                                    if (currentAnswerLines.Count > 0)
                                        currentAnswerLines.Add(line);
                                }

                                j++;
                            }

                            if (currentAnswerLines.Count > 0)
                                answers.Add(string.Join("\n", currentAnswerLines));

                            question.CountAnswer = answers.Count;
                            question.Answers = string.Join("\n\n", answers);

                            if (answers.Count == 1)
                            {
                                string fullAnswer = answers[0].Trim();

                                // Loại bỏ ký hiệu đầu dòng nếu có
                                if (mcAnswerRegex.IsMatch(fullAnswer) || trueFalseAnswer.IsMatch(fullAnswer))
                                {
                                    fullAnswer = fullAnswer.Substring(2).Trim();
                                }

                                question.CorrectAnswer = fullAnswer;

                                if (string.IsNullOrWhiteSpace(fullAnswer))
                                {
                                    question.QuestionType = QuestionType.ShortAnswer;
                                    question.Description += "⚠️ Chưa có nội dung đáp án. ";
                                    question.IsValid = false;
                                }
                                else if (fullAnswer.Length <= 4)
                                {
                                    question.QuestionType = QuestionType.ShortAnswer;
                                    question.IsValid = true;
                                    question.Description += "✅ OK";
                                }
                                else
                                {
                                    question.QuestionType = QuestionType.Essay;
                                    question.IsValid = true;
                                    question.Description += "✅ OK";
                                }
                            }
                            else if (answers.Count == 4 && answers.All(a => mcAnswerRegex.IsMatch(a)))
                            {
                                question.QuestionType = QuestionType.MultipleChoice;
                                question.CorrectAnswer = string.Join(" | ", correctAnswers);

                                bool oneCorrect = correctAnswers.Count == 1;
                                if (!oneCorrect)
                                    question.Description += "⚠️ Phải có đúng 1 đáp án đúng. ";

                                question.IsValid = oneCorrect;
                                if (question.IsValid) question.Description += "✅ OK";
                            }
                            else if (answers.Count == 4 && answers.All(a => trueFalseAnswer.IsMatch(a)))
                            {
                                question.QuestionType = QuestionType.TrueFalse;
                                question.Level = Level.None;
                                question.CorrectAnswer = string.Join(" | ", correctAnswers);
                                question.IsValid = true;
                                question.Description += "✅ OK";
                            }
                            else
                            {
                                question.QuestionType = QuestionType.Unknown;
                                question.Description += "⚠️ Không nhận dạng được dạng câu.";
                                question.IsValid = false;
                            }

                            questions.Add(question);
                            i = j - 1;
                        }
                    }
                }
                catch (IOException ioEx)
                {
                    // lỗi do file đang mở hoặc bị lock
                    MessageHelper.Error($"⚠️ Không thể mở file {filePath}: {ioEx.Message}");
                }
                catch (Exception ex)
                {
                    // các lỗi khác
                    MessageHelper.Error($"❌ Lỗi khi parse file {filePath}: {ex.Message}");
                }

                return questions;
            });
        }

        public async Task<string> ExtractTextAsync(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Không tìm thấy file", filePath);

            return await Task.Run(() =>
            {
                using (var doc = WordprocessingDocument.Open(filePath, false))
                {
                    var body = doc.MainDocumentPart.Document.Body;

                    var lines = new List<string>();

                    foreach (var para in body.Elements<Paragraph>())
                    {
                        string text = para.InnerText.Trim();
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        // Nhiều lựa chọn (A-D)
                        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[A-D][\.:]"))
                        {
                            bool underlined = IsFirstTwoCharsUnderlined(para);
                            if (underlined)
                                text += " (Đúng)";
                        }
                        // Đúng/Sai (a-d)
                        else if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[a-d][\):]"))
                        {
                            bool underlined = IsFirstTwoCharsUnderlined(para);
                            text += underlined ? " (Đúng)" : " (Sai)";
                        }
                        // Trả lời ngắn (A. "??", <= 4 ký tự)
                        else if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[A-D][\.:]\s*[""']?\w{1,4}[""']?$"))
                        {
                            // Giữ nguyên
                        }
                        // Tự luận (A. "*", > 4 ký tự)
                        else if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[A-D][\.:]\s*.+$"))
                        {
                            // Giữ nguyên
                        }

                        lines.Add(text);
                    }

                    return string.Join(Environment.NewLine, lines);
                }
            });
        }

        public async Task AddEndNotesAsync(WordprocessingDocument doc)
        {
            await Task.Run(() =>
            {
                var body = doc.MainDocumentPart.Document.Body;

                // Tạo đoạn văn bản đầu tiên (HẾT) in đậm và căn giữa
                var endNote1 = new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                                              new Run(new Text("-------------------- HẾT --------------------") { Space = SpaceProcessingModeValues.Preserve })
                                              {
                                                  RunProperties = new RunProperties(new Bold())
                                              });

                // Tạo đoạn văn bản thứ hai (Thí sinh không được sử dụng tài liệu) in nghiêng
                var endNote2 = new Paragraph(new Run(new Text("- Thí sinh không được sử dụng tài liệu;") { Space = SpaceProcessingModeValues.Preserve })
                {
                    RunProperties = new RunProperties(new Italic())
                });

                // Tạo đoạn văn bản thứ ba (Giám thị không giải thích gì thêm) in nghiêng
                var endNote3 = new Paragraph(new Run(new Text("- Giám thị không giải thích gì thêm.") { Space = SpaceProcessingModeValues.Preserve })
                {
                    RunProperties = new RunProperties(new Italic())
                });

                // Thêm các đoạn vào cuối body
                body.Append(endNote1);
                body.Append(endNote2);
                body.Append(endNote3);

                doc.MainDocumentPart.Document.Save();
            });
        }

        public async Task AddFooterAsync(WordprocessingDocument doc, string version)
        {
            await Task.Run(() =>
            {
                var footerPart = doc.MainDocumentPart.AddNewPart<FooterPart>();
                var footer = new Footer();

                // Tạo phần thông tin trang
                var paragraph = new Paragraph();

                // Căn phải cho đoạn văn
                var paragraphProperties = new ParagraphProperties();
                paragraphProperties.Append(new Justification() { Val = JustificationValues.Right });
                paragraph.AppendChild(paragraphProperties);

                // Hàm tiện ích tạo Run với font Times New Roman
                Run CreateRun(string text)
                {
                    var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                    var runProps = new RunProperties();
                    runProps.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
                    runProps.Append(new FontSize() { Val = "24" }); // 24 half-points = 12pt
                    run.PrependChild(runProps);
                    return run;
                }

                // Thêm chữ "Trang "
                paragraph.Append(CreateRun("Trang "));

                // Trường cho số trang hiện tại
                var pageNumberField = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                var pageNumber = new Run(new FieldCode("PAGE"));
                var pageNumberEnd = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });

                // Gắn font cho các FieldCode
                foreach (var r in new[] { pageNumber, pageNumberField, pageNumberEnd })
                {
                    var runProps = new RunProperties();
                    runProps.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
                    runProps.Append(new FontSize() { Val = "24" });
                    r.PrependChild(runProps);
                }

                // Trường cho tổng số trang
                var totalPagesField = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                var totalPages = new Run(new FieldCode("SECTIONPAGES"));
                var totalPagesEnd = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });

                foreach (var r in new[] { totalPagesField, totalPages, totalPagesEnd })
                {
                    var runProps = new RunProperties();
                    runProps.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
                    runProps.Append(new FontSize() { Val = "24" });
                    r.PrependChild(runProps);
                }

                // Thêm thông tin vào paragraph
                paragraph.Append(pageNumberField, pageNumber, pageNumberEnd);
                paragraph.Append(CreateRun("/"));
                paragraph.Append(totalPagesField, totalPages, totalPagesEnd);
                paragraph.Append(CreateRun($" - Mã đề {version}"));

                footer.Append(paragraph);
                footerPart.Footer = footer;

                // Thêm footer vào tài liệu
                var sectionProperties = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
                if (sectionProperties != null)
                {
                    var footerReference = new FooterReference()
                    {
                        Id = doc.MainDocumentPart.GetIdOfPart(footerPart),
                        Type = HeaderFooterValues.Default
                    };
                    sectionProperties.Append(footerReference);
                }
                doc.MainDocumentPart.Document.Save();
            });
        }

        public async Task InsertTemplateAsync(string templatePath, WordprocessingDocument doc, MixInfo mixInfo, string code)
        {
            await Task.Run(async () =>
            {
                var targetBody = doc.MainDocumentPart.Document.Body;

                using (var templateDoc = WordprocessingDocument.Open(templatePath, false))
                {
                    var templateBody = templateDoc.MainDocumentPart.Document.Body;

                    // Đồng bộ StyleDefinitionsPart
                    if (templateDoc.MainDocumentPart.StyleDefinitionsPart != null)
                    {
                        if (doc.MainDocumentPart.StyleDefinitionsPart == null)
                        {
                            var stylePart = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                            stylePart.FeedData(templateDoc.MainDocumentPart.StyleDefinitionsPart.GetStream());
                        }
                        else
                        {
                            doc.MainDocumentPart.StyleDefinitionsPart.FeedData(
                                templateDoc.MainDocumentPart.StyleDefinitionsPart.GetStream());
                        }
                    }

                    // Chèn phần tử từ template vào đầu body
                    foreach (var element in templateBody.Elements().Reverse())
                    {
                        targetBody.InsertAt(element.CloneNode(true), 0);
                    }

                    await ReplacePlaceholdersAsync(targetBody, new Dictionary<string, string>
            {
                { "[KYTHI]", mixInfo.TestPeriod ?? string.Empty },
                { "[NAMHOC]", mixInfo.SchoolYear ?? string.Empty },
                { "[MONTHI]", mixInfo.Subject ?? string.Empty },
                { "[DONVICAPTREN]", mixInfo.SuperiorUnit ?? string.Empty },
                { "[DONVI]", mixInfo.Unit ?? string.Empty },
                { "[KHOILOP]", mixInfo.Grade ?? string.Empty },
                { "[MaDe]", code },
                { "[ThoiGian]", mixInfo.Time ?? string.Empty }
            });

                    doc.MainDocumentPart.Document.Save();
                }
            });
        }


        public async Task MoveEssayTableToEndAsync(WordprocessingDocument answerDoc)
        {
            await Task.Run(() =>
            {
                var body = answerDoc.MainDocumentPart?.Document.Body;
                if (body == null) return;

                var tables = body.Elements<Table>().ToList();

                var essayTable = tables.FirstOrDefault(t =>
                    t.Descendants<TableRow>().Any(r => r.InnerText.Contains("Đáp án")) &&
                    t.InnerText.Contains("Câu") &&
                    t.InnerText.Contains("Điểm"));

                if (essayTable != null)
                {
                    // Clone đúng cách bằng XML để giữ nguyên toàn bộ TableProperties
                    string xml = essayTable.OuterXml;
                    var newTable = new Table(xml);

                    essayTable.Remove(); // Xóa bản gốc
                    body.Append(newTable); // Thêm lại bản sao đúng
                }
            });
        }

        public async Task AppendGuideAsync(WordprocessingDocument doc, List<QuestionExport> answers, MixInfo mixInfo, string code)
        {
            await Task.Run(async () =>
            {
                try
                {
                    var mainPart = doc.MainDocumentPart;
                    var document = mainPart.Document;
                    var body = document.Body ?? document.AppendChild(new Body());

                    await ReplacePlaceholdersAsync(body, new Dictionary<string, string>
                    {
                        { "[KYTHI]", mixInfo.TestPeriod ?? string.Empty },
                        { "[NAMHOC]", mixInfo.SchoolYear ?? string.Empty },
                        { "[MONTHI]", mixInfo.Subject ?? string.Empty },
                        { "[KHOILOP]", mixInfo.Grade ?? string.Empty },
                        { "[DONVICAPTREN]", mixInfo.SuperiorUnit ?? string.Empty },
                        { "[DONVI]", mixInfo.Unit ?? string.Empty },
                        { "[MaDe]", code },
                        { "[ThoiGian]", mixInfo.Time ?? string.Empty }
                    });

                    var grouped = answers.GroupBy(a => a.Type).OrderBy(g => g.Key).ToList();

                    int index = 0;
                    string[] points = new string[]
                    {
                        mixInfo.PointMultipleChoice,
                        mixInfo.PointTrueFalse,
                        mixInfo.PointShortAnswer,
                        mixInfo.PointEssay
                    };

                    foreach (var group in grouped)
                    {
                        index++;
                        string title = CreateSectionTitle(group.Key, index, group.Count(), points);

                        if (!string.IsNullOrEmpty(title))
                        {
                            var heading = new Paragraph();
                            var parts = title.Split(new[] { '.' }, 3);
                            if (parts.Length >= 3)
                            {
                                var boldRun = new Run(new RunProperties(new Bold()), new Text($"{parts[0]}.{parts[1]}.") { Space = SpaceProcessingModeValues.Preserve });
                                var normalRun = new Run(new Text(parts[2]) { Space = SpaceProcessingModeValues.Preserve });
                                heading.Append(boldRun, normalRun);
                            }
                            body.Append(heading);
                        }

                        var table = CreateTable();

                        if (group.Key == QuestionType.MultipleChoice)
                        {
                            body.Append(new Paragraph(new Run(new Text("Mỗi câu trả lời đúng thí sinh được 0,25 điểm.") { Space = SpaceProcessingModeValues.Preserve })));
                            AddMultipleChoiceAnswerTable(table, answers.Where(q => q.Type == QuestionType.MultipleChoice).ToList());
                            body.Append(table);
                        }
                        else if (group.Key == QuestionType.TrueFalse)
                        {
                            body.Append(new Paragraph(new Run(new Text("- Thí sinh chỉ lựa chọn chính xác 01 ý trong 01 câu hỏi được 0,1 điểm;") { Space = SpaceProcessingModeValues.Preserve })));
                            body.Append(new Paragraph(new Run(new Text("- Thí sinh chỉ lựa chọn chính xác 02 ý trong 01 câu hỏi được 0,25 điểm;") { Space = SpaceProcessingModeValues.Preserve })));
                            body.Append(new Paragraph(new Run(new Text("- Thí sinh chỉ lựa chọn chính xác 03 ý trong 01 câu hỏi được 0,5 điểm;") { Space = SpaceProcessingModeValues.Preserve })));
                            body.Append(new Paragraph(new Run(new Text("- Thí sinh chỉ lựa chọn chính xác cả 04 ý trong 01 câu hỏi được 1 điểm.") { Space = SpaceProcessingModeValues.Preserve })));
                            AddTrueFalseAnswerTable(table, group.Select(a => a.CorrectAnswer).ToList(), 8);
                            body.Append(table);
                        }
                        else if (group.Key == QuestionType.ShortAnswer)
                        {
                            body.Append(new Paragraph(new Run(new Text($"Mỗi câu trả lời đúng thí sinh được 0,5 điểm.") { Space = SpaceProcessingModeValues.Preserve })));
                            AddShortAnswerTable(table, group.Select(a => a.CorrectAnswer).ToList(), 6);
                            body.Append(table);
                        }
                        else if (group.Key == QuestionType.Essay)
                        {
                            // TODO
                        }

                        mainPart.Document.Save();
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error("DOCX save error: " + ex.Message);
                }
            });
        }

        private QuestionType DetectQuestionType(List<OpenXmlElement> block)
        {
            var paras = block.OfType<Paragraph>().ToList();
            var answers = paras.Where(p => System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[A-D]\.")).ToList();
            var trueFalse = paras.Where(p => System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[a-d]\)")).ToList();

            if (answers.Count >= 2) return QuestionType.MultipleChoice;
            if (trueFalse.Count >= 2) return QuestionType.TrueFalse;
            if (answers.Count == 1)
            {
                var content = answers[0].InnerText.Substring(2).Trim();
                return content.Length > 4 ? QuestionType.Essay : QuestionType.ShortAnswer;
            }

            return QuestionType.ShortAnswer;
        }

        private List<List<OpenXmlElement>> SplitQuestions(Body body)
        {
            var result = new List<List<OpenXmlElement>>();
            var current = new List<OpenXmlElement>();
            System.Text.RegularExpressions.Regex headerRegex = Constants.QuestionHeaderRegex;

            foreach (var el in body.Elements())
            {
                if (el is Paragraph para)
                {
                    var text = para.InnerText.Trim();
                    if (headerRegex.IsMatch(text))
                    {
                        if (current.Count > 0)
                            result.Add(current);
                        current = new List<OpenXmlElement>();
                    }
                }
                current.Add(el.CloneNode(true));
            }

            if (current.Count > 0)
                result.Add(current);

            return result;
        }

        private static bool IsFirstTwoCharsUnderlined(Paragraph para)
        {
            var runs = para.Elements<Run>().ToList();
            int captured = 0;
            bool underlined = false;

            foreach (var run in runs)
            {
                string runText = string.Concat(run.Elements<Text>().Select(t => t.Text));
                if (string.IsNullOrEmpty(runText)) continue;

                foreach (char ch in runText)
                {
                    if (captured < 2)
                    {
                        if (!char.IsWhiteSpace(ch))
                        {
                            captured++;
                            if (run.RunProperties?.Underline != null)
                                underlined = true;
                        }
                    }
                }

                if (captured >= 2) break;
            }

            return underlined;
        }

        private async Task ReplacePlaceholdersAsync(Body body, Dictionary<string, string> replacements)
        {
            await Task.Run(() =>
            {
                foreach (var para in body.Descendants<Paragraph>())
                {
                    UpdateTextPlaceholders(para, replacements);
                    InsertNumPagesField(para);
                }
            });
        }

        private Table CreateTable()
        {
            return new Table(
                new TableProperties(
                    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                    new TableLayout { Type = TableLayoutValues.Autofit },
                    new TableJustification { Val = TableRowAlignmentValues.Center },
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                    )
                )
            );
        }

        private TableCell CreateCell(string text, string width = "1200")
        {
            var lines = text.Split('\n');
            bool isMultiLine = lines.Length > 1;

            var cell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                )
            );

            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = isMultiLine ? JustificationValues.Left : JustificationValues.Center }
                )
            );

            for (int i = 0; i < lines.Length; i++)
            {
                paragraph.Append(new Run(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve }));
                if (i < lines.Length - 1)
                {
                    paragraph.Append(new Break());
                }
            }
            cell.Append(paragraph);
            return cell;
        }

        private TableCell CreateCell(List<OpenXmlElement> elements, string width = "5000")
        {
            var cell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                )
            );

            foreach (var element in elements)
            {
                if (element is Paragraph || element is Table || element is Drawing || element is DocumentFormat.OpenXml.Math.OfficeMath)
                {
                    var cloned = element.CloneNode(true);
                    cell.Append(cloned);
                }
                // Nếu là run đơn lẻ hoặc đoạn trống thì có thể đóng gói lại
                else if (element is Run run)
                {
                    var para = new Paragraph();
                    para.Append(run.CloneNode(true));
                    cell.Append(para);
                }
            }

            return cell;
        }

        private void AddMultipleChoiceAnswerTable(Table table, List<QuestionExport> mcQuestions)
        {
            // ==== 2. Tính số lượng câu và làm tròn lên mốc chẵn 10 ====
            int count = mcQuestions.Count;
            int roundedCount = ((count + 9) / 10) * 10;

            // ==== 3. Tạo từng nhóm 10 câu ====
            for (int i = 0; i < roundedCount; i += 10)
            {
                var numberRow = new TableRow();
                var answerRow = new TableRow();

                // Cột đầu tiên: "Câu" và "Đáp án"
                numberRow.Append(CreateCell("Câu"));
                answerRow.Append(CreateCell("Đáp án"));

                for (int j = i + 1; j <= i + 10; j++)
                {
                    numberRow.Append(CreateCell(j <= count ? j.ToString() : string.Empty));
                    var ans = mcQuestions.FirstOrDefault(q => q.QuestionNumber == j);
                    answerRow.Append(CreateCell(ans?.CorrectAnswer ?? string.Empty));
                }

                table.Append(numberRow);
                table.Append(answerRow);
            }
        }

        private static IEnumerable<T[]> Chunk<T>(IEnumerable<T> source, int size)
        {
            if (size <= 0) throw new ArgumentException("Size must be greater than 0.", nameof(size));

            var bucket = new List<T>(size);

            foreach (var item in source)
            {
                bucket.Add(item);
                if (bucket.Count == size)
                {
                    yield return bucket.ToArray();
                    bucket.Clear();
                }
            }

            if (bucket.Count > 0)
                yield return bucket.ToArray();
        }

        private void AddTrueFalseAnswerTable(Table table, List<string> answers, int maxColumns)
        {
            // Hàng 1: Câu 1 → N
            var headerRow = new TableRow();
            headerRow.Append(CreateCell("Câu"));
            for (int i = 1; i <= maxColumns; i++)
            {
                headerRow.Append(CreateCell(i.ToString()));
            }
            table.Append(headerRow);

            // Hàng 2: Đáp án (gộp a-d vào 1 ô, mỗi dòng là một lựa chọn)
            var answerRow = new TableRow();
            answerRow.Append(CreateCell("Đáp án"));

            for (int i = 0; i < maxColumns; i++)
            {
                if (i < answers.Count)
                {
                    var lines = Chunk(answers[i].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries), 2).Select(p => $"{p[0]} {p[1]}").ToList();
                    string combined = string.Join(Environment.NewLine, lines);
                    answerRow.Append(CreateCell(combined));
                }
                else
                {
                    answerRow.Append(CreateCell(""));
                }
            }
            table.Append(answerRow);
        }

        private void AddShortAnswerTable(Table table, List<string> answers, int maxColumns)
        {
            // Hàng 1: Câu 1 → N
            var headerRow = new TableRow();
            headerRow.Append(CreateCell("Câu"));
            for (int i = 1; i <= maxColumns; i++)
            {
                headerRow.Append(CreateCell(i.ToString()));
            }
            table.Append(headerRow);

            // Hàng 2: Đáp án (gộp a-d vào 1 ô, mỗi dòng là một lựa chọn)
            var answerRow = new TableRow();
            answerRow.Append(CreateCell("Đáp án"));
            for (int i = 0; i < maxColumns; i++)
            {
                answerRow.Append(CreateCell(i < answers.Count ? answers[i] : ""));
            }
            table.Append(answerRow);
        }

        private bool IsUnderlined(Paragraph para, string label)
        {
            foreach (var run in para.Elements<Run>())
            {
                string runText = run.InnerText.Trim();
                if (runText.StartsWith(label, StringComparison.OrdinalIgnoreCase))
                {
                    var underline = run.RunProperties?.Underline;
                    return underline != null && underline.Val != null && underline.Val != UnderlineValues.None;
                }
            }

            return false;
        }

        private Level GetLevelFromText(string value)
        {
            if (value.IndexOf("TH", StringComparison.OrdinalIgnoreCase) >= 0)
                return Level.Understand;

            if (value.IndexOf("VD", StringComparison.OrdinalIgnoreCase) >= 0)
                return Level.Manipulate;

            return Level.Know;
        }

        private string ShortLevelCode(Level level)
        {
            switch (level)
            {
                case Level.Know:
                    return "NB";
                case Level.Understand:
                    return "TH";
                case Level.Manipulate:
                    return "VD";
                default:
                    return string.Empty;
            }
        }

        private void FormatParagraph(Paragraph para, MixInfo mixInfo)
        {
            if (para.ParagraphProperties == null)
            {
                para.ParagraphProperties = new ParagraphProperties();
            }

            // Xóa toàn bộ spacing cũ
            para.ParagraphProperties.RemoveAllChildren<SpacingBetweenLines>();

            // Ép spacing về 0pt trước/sau, line 1.2
            para.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines
            {
                Before = "0",
                After = "0",
                Line = "288", // 1.2 dòng
                LineRule = LineSpacingRuleValues.Auto
            };

            para.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "Normal" };

            // Font Times New Roman 12pt cho tất cả run
            foreach (var run in para.Elements<Run>())
            {
                if (run.RunProperties == null)
                {
                    run.RunProperties = new RunProperties();
                }
                run.RunProperties.RunFonts = new RunFonts { Ascii = mixInfo.FontFamily, HighAnsi = mixInfo.FontFamily };
                double fonSize = Convert.ToDouble(mixInfo.FontSize);
                run.RunProperties.FontSize = new FontSize
                {
                    Val = (fonSize * 2).ToString()
                };
            }
        }

        private bool HasImageOrFormula(List<Paragraph> answerGroup)
        {
            return answerGroup.Any(para =>
                para.Descendants<Drawing>().Any() ||                                // Hình ảnh (Drawing)
                para.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() || // Công thức toán
                para.Descendants<EmbeddedObject>().Any() ||                         // Object nhúng
                para.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Any() ||       // VML Shape
                para.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any()      // VML Image (format cũ)
            );
        }

        private string CreateSectionTitle(QuestionType type, int index, int end, string[] points)
        {
            //return type switch
            //{
            //    QuestionType.MultipleChoice => string.Format(Constants.TITLES[0], Constants.ROMANS[index - 1], end),
            //    QuestionType.TrueFalse => string.Format(Constants.TITLES[1], Constants.ROMANS[index - 1], end),
            //    QuestionType.ShortAnswer => string.Format(Constants.TITLES[2], Constants.ROMANS[index - 1], end),
            //    QuestionType.Essay => string.Format(Constants.TITLES[3], Constants.ROMANS[index - 1], end),
            //    _ => string.Empty
            //};

            int typeIndex = (int)type;
            string point = points[typeIndex];

            // Lấy template gốc (không có {2})
            string baseTitle = string.Format(Constants.TITLES[typeIndex], Constants.ROMANS[index - 1], end);

            // Nếu có điểm thì nối thêm
            if (!string.IsNullOrWhiteSpace(point))
            {
                baseTitle += $" ({point} điểm)";
            }

            return baseTitle;
        }

        private string ExtractPointFromText(string text) => System.Text.RegularExpressions.Regex.Match(text, @"\((\d+[,.]?\d*)\s+điểm\)").Groups[1].Value;

        private void UpdateTextPlaceholders(Paragraph para, Dictionary<string, string> replacements)
        {
            var runs = para.Elements<Run>().ToList();
            if (!runs.Any()) return;

            var fullText = string.Join("", runs.Select(r => string.Concat(r.Elements<Text>().Select(t => t.Text ?? ""))));
            if (!replacements.Keys.Any(k => fullText.Contains(k))) return;

            string modifiedText = fullText;
            foreach (var kvp in replacements)
                modifiedText = modifiedText.Replace(kvp.Key, kvp.Value);

            para.RemoveAllChildren<Run>();
            var newRun = new Run(new Text(modifiedText) { Space = SpaceProcessingModeValues.Preserve });

            if (runs.FirstOrDefault()?.RunProperties != null)
                newRun.RunProperties = (RunProperties)runs.First().RunProperties.CloneNode(true);

            para.AppendChild(newRun);
        }

        private void InsertNumPagesField(Paragraph para)
        {
            var runs = para.Elements<Run>().ToList();
            if (!runs.Any()) return;

            var fullText = string.Join("", runs.Select(r => string.Concat(r.Elements<Text>().Select(t => t.Text ?? ""))));
            if (!fullText.Contains("[NUMPAGES]")) return;

            RunProperties originalProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
            string[] parts = fullText.Split(new[] { "[NUMPAGES]" }, StringSplitOptions.None);

            para.RemoveAllChildren<Run>();

            for (int i = 0; i < parts.Length; i++)
            {
                if (!string.IsNullOrEmpty(parts[i]))
                {
                    var run = new Run(new Text(parts[i]) { Space = SpaceProcessingModeValues.Preserve });
                    if (originalProps != null) run.RunProperties = (RunProperties)originalProps.CloneNode(true);
                    para.AppendChild(run);
                }

                if (i < parts.Length - 1)
                {
                    para.Append(
                        CreateFieldRun(FieldCharValues.Begin, originalProps),
                        CreateCodeRun(" SECTIONPAGES ", originalProps),
                        CreateFieldRun(FieldCharValues.Separate, originalProps),
                        CreateResultRun("1", originalProps),
                        CreateFieldRun(FieldCharValues.End, originalProps)
                    );
                }
            }
        }

        private Run CreateFieldRun(FieldCharValues type, RunProperties props)
        {
            var run = new Run { RunProperties = props?.CloneNode(true) as RunProperties };
            run.AppendChild(new FieldChar { FieldCharType = type });
            return run;
        }

        private Run CreateCodeRun(string code, RunProperties props) =>
            new Run(new FieldCode(code) { Space = SpaceProcessingModeValues.Preserve }) { RunProperties = props?.CloneNode(true) as RunProperties };

        private Run CreateResultRun(string result, RunProperties props) =>
            new Run(new Text(result) { Space = SpaceProcessingModeValues.Preserve }) { RunProperties = props?.CloneNode(true) as RunProperties };

        private async Task UpdateQuestionNumberAsync(Paragraph para, int localQuestion)
        {
            await Task.Run(() =>
            {
                var text = para.Elements<Run>().FirstOrDefault()?.Elements<Text>().FirstOrDefault();
                if (text != null)
                    text.Text = System.Text.RegularExpressions.Regex.Replace(text.Text, @"^Câu\s+\d+", $"Câu {localQuestion}");
            });
        }

        private string UpdateAnswerLabels(List<List<OpenXmlElement>> shuffled, string[] labels)
        {
            string correctAnswer = string.Empty;

            // Duyệt qua từng nhóm đáp án sau khi shuffle
            for (int i = 0; i < shuffled.Count && i < labels.Length; i++)
            {
                var group = shuffled[i];

                // ==============================
                // 1. Xác định đáp án đúng (dựa vào underline)
                // ==============================
                bool isCorrect = group.Any(el =>
                    el.Descendants<Run>().Any(r => r.RunProperties?.Underline?.Val != null &&
                                                   r.RunProperties.Underline.Val != UnderlineValues.None));
                if (isCorrect)
                    correctAnswer = labels[i].Substring(0, 1);

                var firstPara = group.OfType<Paragraph>().FirstOrDefault();
                if (firstPara != null)
                {
                    // ==============================
                    // 2. Lấy tất cả Run gốc để giữ nguyên định dạng
                    // ==============================
                    var runs = firstPara.Elements<Run>().ToList();

                    if (runs.Count > 0)
                    {
                        // Run đầu tiên thường chứa nhãn (A./B./C./D.)
                        var firstRun = runs.First();
                        var textNode = firstRun.GetFirstChild<Text>();

                        if (textNode != null && System.Text.RegularExpressions.Regex.IsMatch(textNode.Text, @"^[A-D]\."))
                        {
                            // ==============================
                            // 3. Thay nhãn trong text, giữ nguyên RunProperties
                            // ==============================
                            textNode.Text = labels[i] + " ";
                        }
                    }

                    // ==============================
                    // 4. Xóa toàn bộ Run cũ và thêm lại các Run đã chỉnh sửa
                    // ==============================
                    firstPara.RemoveAllChildren<Run>();
                    foreach (var run in runs)
                    {
                        firstPara.Append(run.CloneNode(true)); // Clone giữ nguyên RunProperties (sup/sub)
                    }
                }
            }

            return correctAnswer;
        }


        private async Task<List<OpenXmlElement>> ExtractEssayAnswerAsync(List<OpenXmlElement> block)
        {
            return await Task.Run(() =>
            {
                var answerElements = new List<OpenXmlElement>();

                // Tìm paragraph đầu tiên có dạng "A. ..."
                var paras = block.OfType<Paragraph>().ToList();
                var firstAnswerPara = paras.FirstOrDefault(p =>
                    System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"));

                if (firstAnswerPara == null)
                    return answerElements;

                int startIndex = block.IndexOf(firstAnswerPara);

                // ➤ Copy nguyên xi tất cả các phần tử từ đoạn này trở đi
                for (int i = startIndex; i < block.Count; i++)
                {
                    // Nếu gặp một nhãn mới (ví dụ "B. ...") thì dừng
                    if (block[i] is Paragraph p &&
                        System.Text.RegularExpressions.Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+") &&
                        i != startIndex)
                        break;

                    answerElements.Add(block[i]);
                }

                return answerElements;
            });
        }

        private void StripEssayLabel(Paragraph p)
        {
            var runs = p.Elements<Run>().ToList();
            p.RemoveAllChildren<Run>();

            bool removedLetter = false, removedDot = false, removingPrefix = true;

            foreach (var run in runs)
            {
                var clonedRun = (Run)run.CloneNode(true);
                var text = clonedRun.GetFirstChild<Text>();

                if (removingPrefix && text != null && !string.IsNullOrEmpty(text.Text))
                {
                    var s = text.Text.TrimStart();

                    if (!removedLetter && s.Length > 0 && char.IsUpper(s[0]))
                    {
                        s = s.Substring(1);
                        removedLetter = true;
                    }
                    if (removedLetter && !removedDot && s.StartsWith("."))
                    {
                        s = s.Substring(1);
                        removedDot = true;
                    }
                    if (removedLetter && removedDot)
                    {
                        s = s.TrimStart();
                        removingPrefix = false;
                    }

                    text.Text = s;
                    if (string.IsNullOrEmpty(s) && clonedRun.Elements().All(e => e is RunProperties))
                        continue;
                }

                // Bỏ underline và ký hiệu <Đ>
                if (clonedRun.RunProperties?.Underline != null)
                    clonedRun.RunProperties.RemoveAllChildren<Underline>();
                if (text != null && text.Text.Contains("<Đ>"))
                    text.Text = text.Text.Replace("<Đ>", "");

                p.Append(clonedRun);
            }
        }

        private void ProcessVmlElements(OpenXmlElement element, MainDocumentPart sourceMainPart, MainDocumentPart targetMainPart)
        {
            // Xử lý ImageData trong VML shape
            foreach (var vmlShape in element.Descendants<Shape>())
            {
                var imageData = vmlShape.Descendants<ImageData>().FirstOrDefault();
                if (imageData?.RelationshipId?.Value is string vmlRelId)
                {
                    if (sourceMainPart.GetPartById(vmlRelId) is ImagePart sourceVmlImage)
                    {
                        var newVmlImagePart = targetMainPart.AddImagePart(sourceVmlImage.ContentType);
                        using (var stream = sourceVmlImage.GetStream())
                        {
                            newVmlImagePart.FeedData(stream);
                            imageData.RelationshipId.Value = targetMainPart.GetIdOfPart(newVmlImagePart);
                        }
                    }
                }
            }

            // Xử lý group VML đệ quy
            foreach (var vmlGroup in element.Descendants<Group>())
            {
                ProcessVmlElements(vmlGroup, sourceMainPart, targetMainPart);
            }
        }

        private List<OpenXmlElement> CloneAnswerBlock(
      List<OpenXmlElement> answerBlock,
      MainDocumentPart sourcePart,
      MainDocumentPart targetPart)
        {
            var imageRelMap = new Dictionary<string, string>();
            var oleRelMap = new Dictionary<string, string>();
            var chartRelMap = new Dictionary<string, string>();
            var clones = new List<OpenXmlElement>();
            bool isFirstPara = true;

            foreach (var el in answerBlock)
            {
                var clone = el.CloneNode(true);

                // 👉 Nếu là Paragraph thì strip nhãn A./B./C./D. ở đoạn đầu tiên
                if (isFirstPara && clone is Paragraph para)
                {
                    StripEssayLabel(para);
                    isFirstPara = false;
                }

                // 👉 Copy ảnh (Blip)
                foreach (var blip in clone.Descendants<Blip>())
                {
                    var oldRelId = blip.Embed?.Value;
                    if (string.IsNullOrEmpty(oldRelId)) continue;

                    if (!imageRelMap.TryGetValue(oldRelId, out var newRelId))
                    {
                        if (sourcePart.GetPartById(oldRelId) is ImagePart srcImg)
                        {
                            var newImg = targetPart.AddImagePart(srcImg.ContentType);
                            using (var s = srcImg.GetStream())
                                newImg.FeedData(s);

                            newRelId = targetPart.GetIdOfPart(newImg);
                            imageRelMap[oldRelId] = newRelId;
                        }
                    }

                    blip.Embed.Value = imageRelMap[oldRelId];
                }

                // 👉 Copy công thức MathType (OLE Object)
                foreach (var ole in clone.Descendants<OleObject>())
                {
                    var oldRelId = ole.GetAttribute("id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
                    if (string.IsNullOrEmpty(oldRelId)) continue;

                    if (!oleRelMap.TryGetValue(oldRelId, out var newRelId))
                    {
                        var srcOle = sourcePart.GetPartById(oldRelId);
                        var newOle = targetPart.AddEmbeddedPackagePart(srcOle.ContentType);
                        using (var s = srcOle.GetStream())
                            newOle.FeedData(s);

                        newRelId = targetPart.GetIdOfPart(newOle);
                        oleRelMap[oldRelId] = newRelId;
                    }

                    ole.SetAttribute(new OpenXmlAttribute("r", "id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        oleRelMap[oldRelId]));
                }

                // 👉 Copy Chart (nếu có)
                foreach (var chart in clone.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Chart>())
                {
                    var chartRelId = chart.GetAttribute("id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
                    if (string.IsNullOrEmpty(chartRelId)) continue;

                    if (!chartRelMap.TryGetValue(chartRelId, out var newRelId))
                    {
                        var srcChart = sourcePart.GetPartById(chartRelId);
                        var newChart = targetPart.AddPart(srcChart);
                        newRelId = targetPart.GetIdOfPart(newChart);
                        chartRelMap[chartRelId] = newRelId;
                    }

                    chart.SetAttribute(new OpenXmlAttribute("r", "id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        chartRelMap[chartRelId]));
                }

                // 👉 Xử lý VML shapes
                ProcessVmlElements(clone, sourcePart, targetPart);

                clones.Add(clone);
            }

            return clones;
        }

    }
}
