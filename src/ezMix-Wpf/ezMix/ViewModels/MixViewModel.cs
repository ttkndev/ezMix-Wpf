using DocumentFormat.OpenXml.Packaging;
using ezMix.Core;
using ezMix.Helpers;
using ezMix.Models;
using ezMix.Models.Enums;
using ezMix.Services.Interfaces;
using Microsoft.Office.Interop.Word;
using Regex.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace ezMix.ViewModels
{
    public class MixViewModel : ObservableObject
    {
        private readonly IOpenXMLService _openXMLService;
        private readonly IInteropWordService _interopWordService;
        private readonly IGeminiService _geminiService;
        private readonly IExcelService _excelService;

        private readonly string PromptsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Prompts");
        private readonly string PromtAnalyzeExamFile = Path.Combine(Directory.GetCurrentDirectory(), "Prompts", "PromptAnalyzeExam.txt");
        private readonly string ConfigsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Configs");
        private readonly string ConfigsFile = Path.Combine(Directory.GetCurrentDirectory(), "Configs", "Config.xml");

        private int multipleChoiceCount;
        private int trueFalseCount;
        private int shortAnswerCount;
        private int essayCount;
        private int unknownCount;

        private bool hasMultipleChoice = false;
        private bool hasTrueFalse = false;
        private bool hasShortAnswer = false;
        private bool hasEssay = false;
        private bool hasTotalPoint = false;

        private string totalPoint = string.Empty;

        public ObservableCollection<string> FontFamilies { get; } = new ObservableCollection<string>
        {
            "Times New Roman", "Arial", "Tahoma", "Calibri", "Cambria", "Verdana", "Georgia"
        };
        public ObservableCollection<string> FontSizes { get; } = new ObservableCollection<string>
        {
            "10", "11", "12", "13", "14", "16", "18", "20"
        };
        public ProgressOverlay ProgressOverlay { get => progressOverlay; set => SetProperty(ref progressOverlay, value); }
        public string PromptAnalyzeExam { get => promptAnalyzeExam; set => SetProperty(ref promptAnalyzeExam, value); }
        public string SourceFile { get => sourceFile; set => SetProperty(ref sourceFile, value); }
        public string DestinationFile { get => destinationFile; set => SetProperty(ref destinationFile, value); }
        public string OutputFolder { get => outputFolder; set => SetProperty(ref outputFolder, value); }
        public ObservableCollection<Question> Questions { get => questions; set => SetProperty(ref questions, value); }
        public ObservableCollection<ExamType> ExamTypes { get => examTypes; set => SetProperty(ref examTypes, value); }
        public ExamType SelectedExamType { get => selectedExamType; set => SetProperty(ref selectedExamType, value); }
        public MixInfo MixInfo { get => mixInfo; set => SetProperty(ref mixInfo, value); }
        public string ExamCodes { get => examCodes; set => SetProperty(ref examCodes, value); }
        public string ProcessContent { get => processContent; set => SetProperty(ref processContent, value); }
        public string InputText { get => inputText; set => SetProperty(ref inputText, value); }
        public string ResultText { get => resultText; set => SetProperty(ref resultText, value); }
        public bool IsEnableMix { get => isEnableMix; set => SetProperty(ref isEnableMix, value); }
        public bool IsOK { get => isOK; set => SetProperty(ref isOK, value); }
        public int MultipleChoiceCount { get => multipleChoiceCount; set => SetProperty(ref multipleChoiceCount, value); }
        public int TrueFalseCount { get => trueFalseCount; set => SetProperty(ref trueFalseCount, value); }
        public int ShortAnswerCount { get => shortAnswerCount; set => SetProperty(ref shortAnswerCount, value); }
        public int EssayCount { get => essayCount; set => SetProperty(ref essayCount, value); }
        public int UnknownCount { get => unknownCount; set => SetProperty(ref unknownCount, value); }
        public string TemplateFolder { get => templateFolder; set => SetProperty(ref templateFolder, value); }
        public List<string> ShuffledTypes { get => shuffledTypes; set => shuffledTypes = value; }

        private ProgressOverlay progressOverlay = new ProgressOverlay();
        private string promptAnalyzeExam = string.Empty;
        private string sourceFile = string.Empty;
        private string destinationFile = string.Empty;
        private string outputFolder = string.Empty;
        private string templateFolder = Path.Combine(Directory.GetCurrentDirectory(), "Assets", "Templates");
        private ObservableCollection<Question> questions = new ObservableCollection<Question>();
        private ObservableCollection<ExamType> examTypes = new ObservableCollection<ExamType>();
        private ExamType selectedExamType = ExamType.ezMix;
        private MixInfo mixInfo = new MixInfo();
        private string examCodes = string.Empty;
        private string processContent = string.Empty;
        private string inputText = string.Empty;
        private string resultText = string.Empty;
        private bool isEnableMix = false;
        private bool isOK = false;
        private List<string> shuffledTypes = new List<string> { "auto", "equal" };


        public RelayCommand AnalyzeFileAsyncCommand { get; }
        public RelayCommand RecognitionFileAsyncCommand { get; }
        public RelayCommand MixAsyncCommand { get; }
        public RelayCommand SaveConfigCommand { get; }
        public RelayCommand SetConfigCommand { get; }
        public RelayCommand GenerateRandomExamCodesCommand { get; }
        public RelayCommand GenerateSequentialExamCodesCommand { get; }
        public RelayCommand GenerateEvenExamCodesCommand { get; }
        public RelayCommand GenerateOddExamCodesCommand { get; }
        public RelayCommand LoadExamAsyncCommand { get; }
        public RelayCommand AnalyzeByGeminiAsyncCommand { get; }
        public RelayCommand ResetPromptCommand { get; }
        public RelayCommand SavePromptCommand { get; }
        public RelayCommand LoadPdfAndOcrAsyncCommand { get; }
        public RelayCommand LoadImageAndOcrAsyncCommand { get; }
        public RelayCommand OpenResourceCommand { get; }
        public RelayCommand ChangeModelCommand { get; }
        public RelayCommand RefreshStatisticsCommand { get; }

        public bool HasMultipleChoice { get => hasMultipleChoice; set => SetProperty(ref hasMultipleChoice, value); }
        public bool HasTrueFalse { get => hasTrueFalse; set => SetProperty(ref hasTrueFalse, value); }
        public bool HasShortAnswer { get => hasShortAnswer; set => SetProperty(ref hasShortAnswer, value); }
        public bool HasEssay { get => hasEssay; set => SetProperty(ref hasEssay, value); }
        public string TotalPoint { get => totalPoint; set => SetProperty(ref totalPoint, value); }
        public bool HasTotalPoint { get => hasTotalPoint; set => SetProperty(ref hasTotalPoint, value); }

        public MixViewModel(IOpenXMLService openXMLService, IInteropWordService interopWordService, IGeminiService geminiService, IExcelService excelService)
        {
            // Khởi tạo các service
            _openXMLService = openXMLService;
            _interopWordService = interopWordService;
            _geminiService = geminiService;
            _excelService = excelService;
            AddLog("Khởi tạo các service: OpenXML, InteropWord, Excel, Gemini");

            // Khởi tạo các command
            AnalyzeFileAsyncCommand = new RelayCommand(async _ => await AnalyzeFileAsync());
            RecognitionFileAsyncCommand = new RelayCommand(async _ => await RecognitionFileAsync());
            MixAsyncCommand = new RelayCommand(async _ => await MixAsync());
            SaveConfigCommand = new RelayCommand(_ => SaveConfig());
            SetConfigCommand = new RelayCommand(_ => SetConfig());
            GenerateRandomExamCodesCommand = new RelayCommand(_ => GenerateRandomExamCodes());
            GenerateSequentialExamCodesCommand = new RelayCommand(_ => GenerateSequentialExamCodes());
            GenerateEvenExamCodesCommand = new RelayCommand(_ => GenerateEvenExamCodes());
            GenerateOddExamCodesCommand = new RelayCommand(_ => GenerateOddExamCodes());
            LoadExamAsyncCommand = new RelayCommand(async _ => await LoadExamAsync());
            AnalyzeByGeminiAsyncCommand = new RelayCommand(async _ => await AnalyzeByGeminiAsync());
            ResetPromptCommand = new RelayCommand(async _ => await ResetPrompt());
            SavePromptCommand = new RelayCommand(async _ => await SavePrompt());
            LoadPdfAndOcrAsyncCommand = new RelayCommand(async _ => await LoadPdfAndOcrAsync());
            LoadImageAndOcrAsyncCommand = new RelayCommand(async _ => await LoadImageAndOcrAsync());
            OpenResourceCommand = new RelayCommand((param) => OpenResource(param?.ToString() ?? string.Empty));
            ChangeModelCommand = new RelayCommand((param) => ChangeModel(param?.ToString() ?? string.Empty));
            RefreshStatisticsCommand = new RelayCommand(_ => UpdateStatistics());
            AddLog("Khởi tạo các command điều khiển chức năng");

            // Tạo danh sách các loại đề thi từ enum
            ExamTypes = new ObservableCollection<ExamType>(Enum.GetValues(typeof(ExamType)).Cast<ExamType>());
            AddLog("Khởi tạo danh sách ExamTypes từ enum ExamType");

            // Task 1: kiểm tra thư mục Prompts và file phân tích đề
            Task.Run(async () =>
            {
                AddLog("Bắt đầu kiểm tra thư mục Prompts và file phân tích đề");
                if (!Directory.Exists(PromptsFolder))
                {
                    Directory.CreateDirectory(PromptsFolder);
                    AddLog("Tạo mới thư mục PromptsFolder");
                }
                else
                {
                    AddLog("Đã tồn tại thư mục PromptsFolder");
                }

                if (!File.Exists(PromtAnalyzeExamFile))
                {
                    PromptAnalyzeExam = Constants.PromptAnalyzeExam;
                    using (var writer = new StreamWriter(PromtAnalyzeExamFile, false))
                    {
                        await writer.WriteAsync(Constants.PromptAnalyzeExam);
                    }
                    AddLog("Tạo mới file PromtAnalyzeExamFile với nội dung mặc định");
                }
                else
                {
                    using (var reader = new StreamReader(PromtAnalyzeExamFile))
                    {
                        PromptAnalyzeExam = await reader.ReadToEndAsync();
                    }
                    AddLog("Đọc nội dung từ file PromtAnalyzeExamFile");
                }
                AddLog("Hoàn tất kiểm tra thư mục Prompts và file phân tích đề");
            });

            // Task 2: kiểm tra thư mục Configs và file cấu hình MixInfo
            Task.Run(async () =>
            {
                AddLog("Bắt đầu kiểm tra thư mục Configs và file cấu hình MixInfo");
                if (!Directory.Exists(ConfigsFolder))
                {
                    Directory.CreateDirectory(ConfigsFolder);
                    AddLog("Tạo mới thư mục ConfigsFolder");
                }
                else
                {
                    AddLog("Đã tồn tại thư mục ConfigsFolder");
                }

                if (!File.Exists(ConfigsFile))
                {
                    XmlHelper.SaveToXml(ConfigsFile, new MixInfo());
                    AddLog("Tạo mới file ConfigsFile với MixInfo mặc định");
                }
                else
                {
                    MixInfo = XmlHelper.LoadFromXml<MixInfo>(ConfigsFile);
                    AddLog("Đọc cấu hình MixInfo từ file ConfigsFile");
                }
                AddLog("Hoàn tất kiểm tra thư mục Configs và file cấu hình MixInfo");
            });
        }

        private async Task AnalyzeFileAsync()
        {
            try
            {
                var sourcePath = FileHelper.BrowseFile();
                if (string.IsNullOrEmpty(sourcePath))
                    return;

                AddLog("Bắt đầu chức năng chuẩn hóa");
                ShowOverlay("Chuẩn hóa đề kiểm tra", "Đang chuẩn hóa đề kiểm tra...", 0);

                SourceFile = sourcePath;
                AddLog($"Chọn tệp nguồn: {SourceFile}");
                ShowOverlay("Chuẩn hóa đề kiểm tra", "Đã chọn tệp nguồn", 10);

                string sourceFolder = Path.GetDirectoryName(sourcePath);
                string fileName = $"{SelectedExamType}_{Path.GetFileName(sourcePath)}";
                string targetPath = Path.Combine(sourceFolder, fileName);

                if (File.Exists(targetPath))
                {
                    AddLog("Phát hiện tệp ezMix cũ, tiến hành xóa");
                    File.SetAttributes(targetPath, FileAttributes.Normal);
                    File.Delete(targetPath);
                }

                File.Copy(sourcePath, targetPath);
                DestinationFile = targetPath;
                AddLog($"Tạo tệp đích: {DestinationFile}");
                ShowOverlay("Chuẩn hóa đề kiểm tra", "Đã tạo tệp đích", 20);

                AddLog("Bắt đầu chuẩn hóa nội dung...");
                ShowOverlay("Chuẩn hóa đề kiểm tra", "Đang chuẩn hóa nội dung...", 30);
                await ProcessDocumentAsync(
                    DestinationFile,
                    SelectedExamType,
                    MixInfo.IsShowWordWhenAnalyze,
                    msg => ShowOverlay("Chuẩn hóa đề kiểm tra", msg, 40) // callback cập nhật statusText + tiến độ
                );

                AddLog("Phân tích câu hỏi từ tệp đã chuẩn hóa...");
                ShowOverlay("Chuẩn hóa đề kiểm tra", "Đang phân tích câu hỏi...", 80);
                var result = await _openXMLService.ParseDocxQuestionsAsync(DestinationFile);
                Questions = new ObservableCollection<Question>(result);

                // Cập nhật thống kê
                ShowOverlay("Chuẩn hóa đề kiểm tra", "Đang thống kê...", 80);
                AddLog("Thống kê số lượng câu hỏi");
                UpdateStatistics();

                IsOK = Questions.All(q => q.IsValid);
                AddLog(IsOK ? "Tất cả câu hỏi hợp lệ" : "Có câu hỏi không hợp lệ");

                IsEnableMix = !string.IsNullOrEmpty(SourceFile) && File.Exists(SourceFile) && IsOK;
                AddLog(IsEnableMix ? "Cho phép trộn đề" : "Không thể trộn do lỗi");

                MessageHelper.Success($"Chuẩn hóa theo ({SelectedExamType}) thành công");
                AddLog("Chuẩn hóa hoàn tất thành công");

                // Overlay khi hoàn tất
                ShowOverlay("Chuẩn hóa đề kiểm tra", "Chuẩn hóa hoàn tất", 100);
            }
            catch (Exception ex)
            {
                ShowOverlay("Chuẩn hóa đề kiểm tra", $"Lỗi khi chuẩn hóa: {ex.Message}", 0);
                AddLog($"Lỗi khi chuẩn hóa: {ex.Message}");
                MessageHelper.Error(ex);
            }
            finally
            {
                await HideOverlayAsync();
            }
        }

        private async Task RecognitionFileAsync()
        {
            try
            {
                var filePath = FileHelper.BrowseFile();
                if (string.IsNullOrEmpty(filePath))
                    return;

                AddLog("Bắt đầu chức năng nhận dạng");
                ShowOverlay("Nhận dạng đề kiểm tra", "Đang nhận dạng...", 0);

                SourceFile = DestinationFile = filePath;
                AddLog($"Chọn tệp nguồn/đích: {filePath}");
                ShowOverlay("Nhận dạng đề kiểm tra", "Đã chọn tệp", 20);

                AddLog("Phân tích câu hỏi từ file...");
                ShowOverlay("Nhận dạng đề kiểm tra", "Đang phân tích...", 50);
                var result = await _openXMLService.ParseDocxQuestionsAsync(filePath);
                Questions = new ObservableCollection<Question>(result);
                AddLog($"Đã phân tích được {Questions.Count} câu hỏi");

                // Cập nhật thống kê
                ShowOverlay("Nhận dạng đề kiểm tra", "Đang thống kê...", 80);
                AddLog("Thống kê số lượng câu hỏi");
                UpdateStatistics();

                IsOK = Questions.All(q => q.IsValid);
                AddLog(IsOK ? "Tất cả câu hỏi hợp lệ" : "Có câu hỏi không hợp lệ");

                IsEnableMix = File.Exists(SourceFile) && IsOK;
                AddLog(IsEnableMix ? "Tệp hợp lệ, có thể trộn đề" : "Tệp không hợp lệ, không thể trộn đề");

                // Overlay khi hoàn tất
                ShowOverlay("Nhận dạng đề kiểm tra", "Hoàn tất", 100);
            }
            catch (Exception ex)
            {
                ShowOverlay("Nhận dạng đề kiểm tra", $"Lỗi: {ex.Message}", 0);
                AddLog($"Lỗi khi nhận dạng: {ex.Message}");
                MessageHelper.Error(ex);
            }
            finally
            {
                await HideOverlayAsync();
            }
        }

        private async Task MixAsync()
        {
            if (!File.Exists(DestinationFile))
                return;

            try
            {
                AddLog("Bắt đầu chức năng trộn đề");
                ShowOverlay("Trộn đề", "Đang khởi động...", 0);

                OutputFolder = Path.Combine(Path.GetDirectoryName(DestinationFile), "ezMix");

                if (Directory.Exists(OutputFolder))
                {
                    AddLog("Phát hiện thư mục ezMix cũ, tiến hành xóa");
                    ShowOverlay("Trộn đề", "Xóa thư mục ezMix cũ", 10);
                    Directory.Delete(OutputFolder, true);
                }

                Directory.CreateDirectory(OutputFolder);
                AddLog($"Tạo thư mục đầu ra: {OutputFolder}");
                ShowOverlay("Trộn đề", "Đã tạo thư mục đầu ra", 20);

                var versions = ExamCodes.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (versions.Length == 0)
                {
                    ProgressOverlay.IsVisible = false;
                    MessageHelper.Error("Chưa tạo danh sách mã đề");
                    AddLog("Không có mã đề, trộn thất bại");
                    return;
                }

                MixInfo.Versions = versions;
                AddLog($"Danh sách mã đề: {string.Join(", ", versions)}");

                ShowOverlay("Trộn đề", "Đang tạo đề trộn...", 30);
                AddLog("Đang tạo đề trộn...");

                await GenerateShuffledExamsAsync(
                    DestinationFile,
                    OutputFolder,
                    MixInfo,
                    (msg, progress) => ShowOverlay("Trộn đề", msg, progress)
                );

                ShowOverlay("Trộn đề", "Hoàn tất", 100);
                AddLog("Trộn đề hoàn tất");
                MessageHelper.Success("Trộn đề hoàn tất");

                if (Directory.Exists(OutputFolder))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = OutputFolder,
                        UseShellExecute = true
                    });
                    AddLog("Mở thư mục kết quả");
                }
            }
            catch (Exception ex)
            {
                ShowOverlay("Trộn đề", $"Lỗi: {ex.Message}", 0);
                AddLog($"Lỗi khi trộn đề: {ex.Message}");
                MessageHelper.Error(ex);
            }
            finally
            {
                await HideOverlayAsync();
            }
        }

        private void SaveConfig()
        {
            try
            {
                AddLog("Bắt đầu lưu cấu hình");

                // 💾 Lưu đối tượng MixInfo vào file XML cấu hình
                XmlHelper.SaveToXml(ConfigsFile, MixInfo);
                AddLog($"Đã lưu MixInfo vào file cấu hình: {ConfigsFile}");

                // ✅ Thông báo lưu thành công
                MessageHelper.Success("Đã lưu thông tin cấu hình");
                AddLog("Lưu cấu hình hoàn tất thành công");
            }
            catch (Exception ex)
            {
                // ❌ Báo lỗi nếu có sự cố khi lưu
                MessageHelper.Error($"Lỗi khi lưu cấu hình: {ex.Message}");
                AddLog($"Lỗi khi lưu cấu hình: {ex.Message}");
            }
        }

        private void SetConfig()
        {
            try
            {
                AddLog("Bắt đầu nạp lại cấu hình mặc định");

                var dialog = MessageHelper.Question(
                    "Bạn có chắc chắn muốn đặt lại cấu hình mặc định không?",
                    "Xác nhận",
                    System.Windows.MessageBoxImage.Question);

                if (dialog == System.Windows.MessageBoxResult.No)
                {
                    AddLog("Người dùng đã hủy thao tác nạp lại cấu hình");
                    return;
                }

                MixInfo = new MixInfo();
                AddLog("Đã đặt lại cấu hình mặc định");

                XmlHelper.SaveToXml(ConfigsFile, MixInfo);
                AddLog($"Đã lưu cấu hình mặc định vào file: {ConfigsFile}");

                Questions = new ObservableCollection<Question>();
                SourceFile  = DestinationFile = string.Empty;
                UpdateStatistics();

                MessageHelper.Success("Đã nạp lại cấu hình mặc định");
                AddLog("Nạp lại cấu hình mặc định hoàn tất thành công");
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi khi nạp cấu hình: {ex.Message}");
                AddLog($"Lỗi khi nạp cấu hình: {ex.Message}");
            }
        }

        private void GenerateRandomExamCodes()
        {
            try
            {
                AddLog("Bắt đầu sinh mã đề ngẫu nhiên");

                var codes = new HashSet<string>();
                Random random = new Random();

                codes.Add("000");
                AddLog("Đã thêm mã mặc định: 000");

                int prefix = 1;
                while (codes.Count < MixInfo.NumberOfVersions + 1) // +1 vì có "000"
                {
                    string code;
                    do
                    {
                        code = $"{prefix}{random.Next(100):D2}";
                    } while (codes.Contains(code));

                    codes.Add(code);
                    AddLog($"Đã sinh mã: {code}");

                    // tăng prefix, quay lại 1 nếu >9
                    prefix++;
                    if (prefix > 9) prefix = 1;
                }

                ExamCodes = string.Join(" ", codes.OrderBy(c => c));
                AddLog($"Danh sách mã đề: {ExamCodes}");
                AddLog("Sinh mã đề ngẫu nhiên hoàn tất");
            }
            catch (Exception ex)
            {
                AddLog($"Lỗi khi sinh mã đề ngẫu nhiên: {ex.Message}");
                MessageHelper.Error(ex);
            }
        }

        private void GenerateSequentialExamCodes()
        {
            try
            {
                AddLog("Bắt đầu sinh mã đề tuần tự");

                string prefix = string.IsNullOrWhiteSpace(MixInfo.StartCode) ? "00" : MixInfo.StartCode.Trim();
                AddLog($"Prefix sử dụng: {prefix}");

                if (MixInfo.NumberOfVersions > 99)
                {
                    AddLog("Cảnh báo: Không thể sinh quá 99 mã tuần tự (01–99).");
                    MessageHelper.Error("Số lượng đề yêu cầu vượt quá 99. Vui lòng giảm số lượng hoặc chọn cách sinh khác.");
                    return;
                }

                var codes = new List<string> { "000" }
                    .Concat(Enumerable.Range(1, MixInfo.NumberOfVersions)
                    .Select(i => $"{prefix}{i:D2}"));

                ExamCodes = string.Join(" ", codes);

                AddLog($"Danh sách mã đề tuần tự: {ExamCodes}");
                AddLog("Sinh mã đề tuần tự hoàn tất");
            }
            catch (Exception ex)
            {
                AddLog($"Lỗi khi sinh mã đề tuần tự: {ex.Message}");
                MessageHelper.Error(ex);
            }
        }

        private void GenerateEvenExamCodes()
        {
            try
            {
                AddLog("Bắt đầu sinh mã đề chẵn liên tục");

                // Lấy prefix từ MixInfo.StartCode, nếu rỗng thì mặc định "00"
                string prefix = string.IsNullOrWhiteSpace(MixInfo.StartCode) ? "00" : MixInfo.StartCode.Trim();
                AddLog($"Prefix sử dụng: {prefix}");

                // Kiểm tra số lượng đề
                if (MixInfo.NumberOfVersions > 50)
                {
                    AddLog("Cảnh báo: Không thể sinh quá 50 mã chẵn (00–98).");
                    MessageHelper.Error("Số lượng đề yêu cầu vượt quá 50. Vui lòng giảm số lượng hoặc chọn cách sinh khác.");
                    return;
                }

                // Mã mặc định
                var codes = new List<string> { "000" };

                // Sinh các mã chẵn liên tục
                for (int i = 0; i < MixInfo.NumberOfVersions; i++)
                {
                    int evenNumber = i * 2; // 0, 2, 4, 6, ...
                    string code = $"{prefix}{evenNumber:D2}";
                    codes.Add(code);
                    AddLog($"Đã sinh mã chẵn: {code}");
                }

                ExamCodes = string.Join(" ", codes);
                AddLog($"Danh sách mã đề chẵn: {ExamCodes}");
                AddLog("Sinh mã đề chẵn liên tục hoàn tất");
            }
            catch (Exception ex)
            {
                AddLog($"Lỗi khi sinh mã đề chẵn liên tục: {ex.Message}");
                MessageHelper.Error(ex);
            }
        }

        private void GenerateOddExamCodes()
        {
            try
            {
                AddLog("Bắt đầu sinh mã đề lẻ liên tục");

                // Lấy prefix từ MixInfo.StartCode, nếu rỗng thì mặc định "00"
                string prefix = string.IsNullOrWhiteSpace(MixInfo.StartCode) ? "00" : MixInfo.StartCode.Trim();
                AddLog($"Prefix sử dụng: {prefix}");

                // Kiểm tra số lượng đề
                if (MixInfo.NumberOfVersions > 50)
                {
                    AddLog("Cảnh báo: Không thể sinh quá 50 mã lẻ (01–99).");
                    MessageHelper.Error("Số lượng đề yêu cầu vượt quá 50. Vui lòng giảm số lượng hoặc chọn cách sinh khác.");
                    return;
                }

                // Mã mặc định
                var codes = new List<string> { "000" };

                // Sinh các mã lẻ liên tục
                for (int i = 0; i < MixInfo.NumberOfVersions; i++)
                {
                    int oddNumber = i * 2 + 1; // 1, 3, 5, …, 99
                    string code = $"{prefix}{oddNumber:D2}";
                    codes.Add(code);
                    AddLog($"Đã sinh mã lẻ: {code}");
                }

                ExamCodes = string.Join(" ", codes);
                AddLog($"Danh sách mã đề lẻ: {ExamCodes}");
                AddLog("Sinh mã đề lẻ liên tục hoàn tất");
            }
            catch (Exception ex)
            {
                AddLog($"Lỗi khi sinh mã đề lẻ liên tục: {ex.Message}");
                MessageHelper.Error(ex);
            }
        }


        private async Task LoadExamAsync()
        {
            try
            {
                if (File.Exists(DestinationFile))
                {
                    AddLog("Bắt đầu tải đề kiểm tra từ tệp đích");
                    AddLog($"Đã phát hiện tệp: {DestinationFile}");

                    InputText = await _openXMLService.ExtractTextAsync(DestinationFile);

                    AddLog("Đã trích xuất văn bản từ tệp thành công");
                }
                else
                {
                    AddLog("Không tìm thấy tệp đích để tải đề kiểm tra");
                }
            }
            catch (Exception ex)
            {
                AddLog($"Lỗi khi tải đề kiểm tra: {ex.Message}");
                MessageHelper.Error(ex);
            }
        }

        private async Task AnalyzeByGeminiAsync()
        {
            if (string.IsNullOrWhiteSpace(InputText))
                return;

            try
            {
                AddLog("Bắt đầu phân tích đề bằng Gemini");
                ShowOverlay("Phân tích đề bằng Gemini", "Đang phân tích...", 0);

                ResultText = "Đang phân tích...";
                string promptAnalyzeExam = string.Format(PromptAnalyzeExam, MixInfo.Subject, MixInfo.Grade);
                string prompt = $"{promptAnalyzeExam}\n\nĐỀ KIỂM TRA:\n{InputText}";

                // Gọi API Gemini
                ShowOverlay("Phân tích đề bằng Gemini", "Đang gọi Gemini...", 50);
                ResultText = await _geminiService.CallGeminiAsync(MixInfo.GeminiModel, MixInfo.GeminiApiKey, prompt);

                // Overlay khi hoàn tất
                ShowOverlay("Phân tích đề bằng Gemini", "Hoàn tất", 100);
                AddLog("Phân tích đề bằng Gemini hoàn tất");
            }
            catch (Exception ex)
            {
                ShowOverlay("Phân tích đề bằng Gemini", $"Lỗi: {ex.Message}", 0);
                AddLog($"Lỗi khi phân tích bằng Gemini: {ex.Message}");
                ResultText = $"Lỗi khi phân tích: {ex.Message}";
            }
            finally
            {
                await HideOverlayAsync();
            }
        }

        private async Task ResetPrompt()
        {
            try
            {
                AddLog("Bắt đầu reset PromptAnalyzeExam về mặc định");

                PromptAnalyzeExam = Constants.PromptAnalyzeExam;
                AddLog("Đã gán PromptAnalyzeExam từ Constants");

                if (!Directory.Exists(PromptsFolder))
                {
                    Directory.CreateDirectory(PromptsFolder);
                    AddLog($"Tạo thư mục PromptsFolder: {PromptsFolder}");
                }

                using (var writer = new StreamWriter(PromtAnalyzeExamFile, false))
                {
                    await writer.WriteAsync(PromptAnalyzeExam);
                    AddLog($"Đã ghi PromptAnalyzeExam vào file: {PromtAnalyzeExamFile}");
                }

                MessageHelper.Success("✅ PromptAnalyzeExam được reset về mặc định");
                AddLog("Reset PromptAnalyzeExam hoàn tất thành công");
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"❌ Lỗi khi reset PromptAnalyzeExam: {ex.Message}");
                AddLog($"Lỗi khi reset PromptAnalyzeExam: {ex.Message}");
            }
        }

        private async Task SavePrompt()
        {
            try
            {
                AddLog("Bắt đầu lưu PromptAnalyzeExam");

                if (!Directory.Exists(PromptsFolder))
                {
                    Directory.CreateDirectory(PromptsFolder);
                    AddLog($"Tạo thư mục PromptsFolder: {PromptsFolder}");
                }

                using (var writer = new StreamWriter(PromtAnalyzeExamFile, false))
                {
                    await writer.WriteAsync(Constants.PromptAnalyzeExam);
                    AddLog($"Đã ghi PromptAnalyzeExam vào file: {PromtAnalyzeExamFile}");
                }

                MessageHelper.Success("💾 PromptAnalyzeExam đã được lưu thành công");
                AddLog("Lưu PromptAnalyzeExam hoàn tất thành công");
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"❌ Lỗi khi lưu PromptAnalyzeExam: {ex.Message}");
                AddLog($"Lỗi khi lưu PromptAnalyzeExam: {ex.Message}");
            }
        }

        private async Task LoadPdfAndOcrAsync()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(MixInfo.GeminiModel) || string.IsNullOrWhiteSpace(MixInfo.GeminiApiKey))
                {
                    InputText += "\nChưa nhập Gemini Model hoặc Gemini API Key, không thể chạy.";
                    AddLog("Thiếu Gemini Model hoặc API Key, không thể xử lý PDF");
                    return;
                }

                var path = FileHelper.BrowsePdf();
                if (!string.IsNullOrEmpty(path))
                {
                    AddLog("Bắt đầu trích xuất văn bản từ PDF");
                    AddLog($"Đã chọn file: {path}");

                    // Overlay mở đầu
                    ShowOverlay("Trích xuất PDF", "Đang trích xuất...", 0);

                    InputText = $"Đã chọn: {path}\n\n";
                    InputText += "Đang trích xuất văn bản từ PDF...\n";

                    // Gọi Gemini để OCR
                    ShowOverlay("Trích xuất PDF", "Đang gọi Gemini...", 50);
                    var result = await ExtractTextByGeminiAsync(MixInfo.GeminiModel, MixInfo.GeminiApiKey, path);
                    InputText += result;

                    // Overlay khi hoàn tất
                    ShowOverlay("Trích xuất PDF", "Hoàn tất", 100);
                    AddLog("Trích xuất văn bản từ PDF hoàn tất");
                }
            }
            catch (Exception ex)
            {
                ShowOverlay("Trích xuất PDF", $"Lỗi: {ex.Message}", 0);
                AddLog($"Lỗi khi xử lý PDF: {ex.Message}");
                InputText += $"\nLỗi khi xử lý PDF: {ex.Message}";
            }
            finally
            {
                await HideOverlayAsync();
            }
        }

        private async Task LoadImageAndOcrAsync()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(MixInfo.GeminiModel) || string.IsNullOrWhiteSpace(MixInfo.GeminiApiKey))
                {
                    InputText += "\nChưa nhập Gemini Model hoặc Gemini API Key, không thể chạy.";
                    AddLog("Thiếu Gemini Model hoặc API Key, không thể xử lý ảnh");
                    return;
                }

                var path = FileHelper.BrowseImage();
                if (!string.IsNullOrEmpty(path))
                {
                    AddLog("Bắt đầu trích xuất văn bản từ ảnh");
                    AddLog($"Đã chọn file: {path}");

                    // Overlay mở đầu
                    ShowOverlay("Trích xuất ảnh", "Đang trích xuất...", 0);

                    InputText = $"Đã chọn: {path}\n\n";
                    InputText += "Đang trích xuất văn bản từ ảnh...\n";

                    // Gọi Gemini để OCR
                    ShowOverlay("Trích xuất ảnh", "Đang gọi Gemini...", 50);
                    var result = await ExtractTextByGeminiAsync(MixInfo.GeminiModel, MixInfo.GeminiApiKey, path);
                    InputText += result;

                    // Overlay khi hoàn tất
                    ShowOverlay("Trích xuất ảnh", "Hoàn tất", 100);
                    AddLog("Trích xuất văn bản từ ảnh hoàn tất");
                }
            }
            catch (Exception ex)
            {
                ShowOverlay("Trích xuất ảnh", $"Lỗi: {ex.Message}", 0);
                AddLog($"Lỗi khi xử lý ảnh: {ex.Message}");
                InputText += $"\nLỗi khi xử lý ảnh: {ex.Message}";
            }
            finally
            {
                await HideOverlayAsync();
            }
        }

        private void OpenResource(string path)
        {
            FileHelper.OpenResource(path);
        }

        private void ChangeModel(string model)
        {
            if (!string.IsNullOrWhiteSpace(model))
            {
                if (MixInfo != null)
                {
                    MixInfo.GeminiModel = model;
                }
            }
        }

        private void AddLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logLine = $"[{timestamp}] {message}";

            if (string.IsNullOrWhiteSpace(ProcessContent))
            {
                ProcessContent = logLine;
            }
            else
            {
                ProcessContent = $"{ProcessContent}{Environment.NewLine}{logLine}";
            }
        }

        public string ParseGeminiResponse(string jsonResponse)
        {
            using (var doc = JsonDocument.Parse(jsonResponse))
            {
                var root = doc.RootElement;

                var text = root
                    .GetProperty("candidates")[0]
                    .GetProperty("content")
                    .GetProperty("parts")[0]
                    .GetProperty("text")
                    .GetString();

                return text;
            }
        }

        private async Task<string> ExtractTextByGeminiAsync(string model, string apiKey, string path)
        {
            try
            {
                string ext = Path.GetExtension(path).ToLowerInvariant();
                string text = string.Empty;

                switch (ext)
                {
                    case ".pdf":
                        text = await _geminiService.CallGeminiExtractTextFromPdfAsync(model, apiKey, path, null);
                        break;

                    case ".png":
                    case ".jpg":
                    case ".jpeg":
                    case ".bmp":
                        text = await _geminiService.CallGeminiExtractTextFromImageAsync(model, apiKey, path, null);
                        break;

                    default:
                        MessageHelper.Error("Định dạng tệp không được hỗ trợ");
                        return string.Empty;
                }

                return ParseGeminiResponse(text);
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi khi xử lý OCR: {ex.Message}");
                return string.Empty;
            }
        }

        private async Task ProcessDocumentAsync(string filePath, ExamType typeExam, bool isShowWordWhenAnalyze, Action<string> updateOverlay)
        {
            _Document document = null;
            try
            {
                updateOverlay("Đang mở tài liệu Word...");
                AddLog("Mở tài liệu Word");
                document = await _interopWordService.OpenDocumentAsync(filePath, visible: isShowWordWhenAnalyze);
                document.Activate();

                updateOverlay("Đang định dạng lại tài liệu (xóa header/footer, chuyển list thành text, bỏ track changes)...");
                AddLog("Định dạng lại tài liệu (xóa header/footer, chuyển list thành text, bỏ track changes)"); await _interopWordService.FormatDocumentAsync(document);
                await _interopWordService.DeleteAllHeadersAndFootersAsync(document);
                await _interopWordService.ConvertListFormatToTextAsync(document);
                await _interopWordService.RejectAllChangesAsync(document);

                updateOverlay("Thay thế các ký tự thừa");
                AddLog("Thay thế các ký tự thừa");
                var fixs = new Dictionary<string, string>
                {
                    ["^p "] = "^p",     // Sau dấu Enter có dấu cách
                    [" ^p"] = "^p",     // Trước dấu Enter có dấu cách
                    ["  "] = " ",
                    [" ?"] = "?",
                    [" ."] = ".",
                    ["?."] = "?",
                };
                await _interopWordService.ReplaceUntilDoneAsync(document, fixs, matchCase: true, matchWholeWord: false, matchWildcards: false);

                updateOverlay("Chuẩn hóa ký hiệu câu hỏi và đáp án");
                AddLog("Chuẩn hóa ký hiệu câu hỏi và đáp án");
                var replacements = new Dictionary<string, string>
                {
                    ["^t"] = " ",
                    ["^l"] = " ",
                    ["^s"] = " ",
                    ["<$>"] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["A. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["B. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["C. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["D. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["Đáp án: "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["Đáp án. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["ĐÁP ÁN: "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["ĐÁP ÁN. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["ĐÁ:"] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["ĐÁ."] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["ĐA:"] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["ĐA."] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["<#>"] = Constants.QUESTION_TEMPLATE,
                    //["#"] = Constants.QUESTION_TEMPLATE, có thể ảnh hưởng nộ dung
                    ["[<br>]"] = Constants.QUESTION_TEMPLATE,
                    ["<NB>"] = Constants.QUESTION_TEMPLATE,
                    ["<TH>"] = Constants.QUESTION_TEMPLATE,
                    ["<VD>"] = Constants.QUESTION_TEMPLATE,
                    ["<VDC>"] = Constants.QUESTION_TEMPLATE,
                    ["<Đ>"] = "a) ",
                    ["<S>"] = "a) "
                };
                await _interopWordService.ReplaceAsync(document, replacements, matchCase: true, matchWholeWord: false);

                updateOverlay("Thay thế chữ màu đỏ bằng chữ gạch chân");
                AddLog("Thay thế chữ màu đỏ bằng chữ gạch chân");
                await _interopWordService.ReplaceRedTextWithUnderlineAsync(document);

                updateOverlay("Thiết lập phông chữ, màu chữ và cỡ chữ mặc định");
                AddLog("Thiết lập phông chữ, màu chữ và cỡ chữ mặc định");
                var range = document.Range();
                range.Font.Color = WdColor.wdColorBlack;
                range.Font.Name = MixInfo.FontFamily;
                range.Font.Size = Convert.ToSingle(MixInfo.FontSize);

                var removeStarts = new[]
                {
                    "phần 1", "phần 2", "phần 3", "phần 4",
                    "phần i", "phần ii", "phần iii", "phần iv",
                    "dạng 1", "dạng 2", "dạng 3", "dạng 4",
                    "dạng i", "dạng ii", "dạng iii", "dạng iv",
                    "i.", "ii.", "iii.", "iv.",
                    "<g0>", "<g1>", "<g2>", "<g3>",
                    "<#g0>", "<#g1>", "<#g2>", "<#g3>",
                    "---HẾT", "---", "- Thí sinh không", "- Giám thị không", "(Thí sinh không", "(Giám thị không"
                };

                var questionPatterns = new[]
                {
                    /*"Câu [0-9]{1,4} ",*/ "Câu [0-9]{1,4}:", "Câu [0-9]{1,4}.",
                    //"Câu ? ", "Câu ?? ", "Câu ??? ",
                    "Câu ?:", "Câu ??:", "Câu ???:",
                    "Câu ?.", "Câu ??.", "Câu ???."
                };

                updateOverlay("Đang chuẩn hóa các Paragraph...");
                AddLog("Bắt đầu chuẩn hóa các Paragraph...");
                foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in document.Paragraphs)
                {
                    //paragraph.set_Style("Normal");

                    string str = paragraph.Range.Text.Trim();

                    var rangeParagraph = paragraph.Range;
                    rangeParagraph.Font.Name = MixInfo.FontFamily;
                    rangeParagraph.Font.Size = Convert.ToSingle(MixInfo.FontSize);
                    var format = rangeParagraph.ParagraphFormat;

                    rangeParagraph.ListFormat.RemoveNumbers();
                    format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
                    format.TabStops.ClearAll();
                    format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    format.LeftIndent = format.RightIndent = format.FirstLineIndent = 0f;
                    format.SpaceBefore = format.SpaceAfter = 0f;
                    format.KeepWithNext = format.KeepTogether = 0;
                    format.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                    format.LineSpacing = 14.4f;

                    await _interopWordService.ClearTabStopsAsync(paragraph);

                    if (string.IsNullOrEmpty(str) || str.Equals(Constants.QUESTION_TEMPLATE) || str.Equals(Constants.ANSWER_TEMPLATE) ||
                        removeStarts.Any(prefix => str.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                    {
                        paragraph.Range.Delete();
                        continue;
                    }

                    if (str.StartsWith("Câu", StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var pattern in questionPatterns)
                        {
                            await _interopWordService.ReplaceFirstAsync(paragraph, pattern, Constants.QUESTION_TEMPLATE, matchWildcards: true);
                        }
                    }

                    var match = System.Text.RegularExpressions.Regex.Match(str, @"^\s*([a-d])[\.\)]");
                    if (match.Success)
                    {
                        var label = match.Groups[1].Value + ") ";
                        await _interopWordService.ReplaceFirstAsync(paragraph, match.Value.Trim(), label);
                    }

                    if (typeExam == ExamType.Intest || typeExam == ExamType.MasterTest)
                    {
                        var matchTF = System.Text.RegularExpressions.Regex.Match(str, @"^([a-d])\)");
                        if (matchTF.Success)
                        {
                            var rangeTF = paragraph.Range;
                            bool isUnderlined = rangeTF.Characters[1].Font.Underline == Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                                             && rangeTF.Characters[2].Font.Underline == Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                            string replacement = isUnderlined ? "<Đ>" : "<S>";
                            await _interopWordService.ReplaceFirstAsync(paragraph, matchTF.Value.Trim(), replacement);
                        }
                    }

                    if (paragraph.Range.InlineShapes.Count == 1 && str == "/")
                    {
                        paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }

                updateOverlay("Hoàn tất chuẩn hóa các Paragraph");
                AddLog("Hoàn tất chuẩn hóa các Paragraph"); 

                updateOverlay("Chuẩn hóa ký hiệu câu hỏi theo loại đề");
                AddLog("Chuẩn hóa ký hiệu câu hỏi theo loại đề");
                string symbolQuestion;
                switch (typeExam)
                {
                    case ExamType.MasterTest:
                    case ExamType.Intest:
                        symbolQuestion = "<#>";
                        break;
                    case ExamType.MCMix:
                        symbolQuestion = "[<br>]";
                        break;
                    //case ExamType.SmartTest:
                    //    symbolQuestion = "#";
                    //    break;
                    default:
                        symbolQuestion = string.Empty;
                        break;
                }

                if (string.IsNullOrEmpty(symbolQuestion))
                {
                    await _interopWordService.SetQuestionsToNumberAsync(document);
                }
                else
                {
                    await _interopWordService.ReplaceAsync(document, new Dictionary<string, string>
                    {
                        [Constants.QUESTION_TEMPLATE] = symbolQuestion
                    }, true);
                }

                updateOverlay("Chuẩn hóa ký hiệu đáp án");
                AddLog("Chuẩn hóa ký hiệu đáp án");
                if (typeExam == ExamType.MasterTest || typeExam == ExamType.Intest)
                {
                    await _interopWordService.ReplaceAsync(document, new Dictionary<string, string>
                    {
                        [Constants.ANSWER_TEMPLATE] = "<$>"
                    }, true);
                    await _interopWordService.ReplaceUnderlineWithRedTextAsync(document);
                }
                else
                {
                    await _interopWordService.SetAnswersToABCDAsync(document);
                }

                updateOverlay("Xử lý các thay thế cuối cùng, sửa lỗi MathType nếu cần");
                AddLog("Xử lý các thay thế cuối cùng, sửa lỗi MathType nếu cần");
                await _interopWordService.ReplaceUntilDoneAsync(document, new Dictionary<string, string>
                {
                    ["^p "] = "^p",
                    [" ^p"] = "^p",
                    ["  "] = " ",
                    ["<#> "] = "<#>",
                    ["<Đ> "] = "<Đ>",
                    ["<S> "] = "<S>",
                });

                if (MixInfo.IsFixMathType)
                {
                    await _interopWordService.FixMathTypeAsync(document);
                }

                _interopWordService.NormalizeParagraphEnds(document);
                await _interopWordService.FormatQuestionAndAnswerAsync(document);
                await _interopWordService.SaveDocumentAsync(document);

                updateOverlay("Hoàn tất chuẩn hóa tài liệu");
                AddLog("Hoàn tất chuẩn hóa tài liệu");
            }
            catch (Exception ex)
            {
                AddLog($"Lỗi khi chuẩn hóa tài liệu: {ex.Message}");
                MessageHelper.Error(ex);
            }
            finally
            {
                if (document != null)
                {
                    await _interopWordService.CloseDocumentAsync(document);
                    await _interopWordService.QuitWordAppAsync();
                }
            }
        }

        private async Task GenerateShuffledExamsAsync(string sourceFile, string outputFolder, MixInfo mixInfo, Action<string, double> updateOverlay)
        {
            await Task.Run(async () =>
            {
                string mixTemplate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Constants.TEMPLATES_FOLDER, Constants.MIX_TEMPLATE_FILE);
                string guideTemplate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Constants.TEMPLATES_FOLDER, Constants.GUIDE_TEMPLATE_FILE);

                int totalVersions = mixInfo.Versions.Length;
                double baseProgress = 10;   // sau bước chuẩn bị thư mục
                double stepPerVersion = 70.0 / totalVersions; // chia đều 70% cho các phiên bản

                // Chuẩn bị thư mục
                if (Directory.Exists(outputFolder))
                {
                    Directory.Delete(outputFolder, true);
                    AddLog("Đã xóa thư mục output cũ");
                    updateOverlay?.Invoke("Xóa thư mục output cũ", 5);
                }
                Directory.CreateDirectory(outputFolder);
                AddLog($"Tạo thư mục output: {outputFolder}");
                updateOverlay?.Invoke("Tạo thư mục output", 10);

                string answerFolder = Path.Combine(outputFolder, Constants.ANSWERS_FOLER);
                Directory.CreateDirectory(answerFolder);
                AddLog($"Tạo thư mục đáp án: {answerFolder}");
                updateOverlay?.Invoke("Tạo thư mục đáp án", 15);

                var allAnswers = new List<QuestionExport>();

                // Vòng lặp qua từng mã đề
                int index = 0;
                foreach (var version in mixInfo.Versions)
                {
                    index++;
                    double versionStart = baseProgress + (index - 1) * stepPerVersion;
                    double versionEnd = baseProgress + index * stepPerVersion;

                    AddLog($"Bắt đầu xử lý mã đề {version}");
                    updateOverlay?.Invoke($"Đang xử lý mã đề {version}", versionStart);

                    string mixFile = Path.Combine(outputFolder, $"{Constants.EXAM_PREFIX}{version}.docx");
                    File.Copy(sourceFile, mixFile, true);
                    AddLog($"Tạo file đề: {mixFile}");

                    string answerFile = Path.Combine(answerFolder, $"{Constants.ANSWER_PREFIX}{version}.docx");
                    File.Copy(guideTemplate, answerFile, true);
                    AddLog($"Tạo file đáp án: {answerFile}");

                    using (var mixDoc = WordprocessingDocument.Open(mixFile, true))
                    using (var answerDoc = WordprocessingDocument.Open(answerFile, true))
                    {
                        updateOverlay?.Invoke($"Shuffle câu hỏi (mã đề {version})", versionStart + stepPerVersion * 0.2);
                        var answers = await _openXMLService.ShuffleQuestionsAsync(mixDoc, version, answerDoc, mixInfo);
                        AddLog($"Shuffle câu hỏi xong cho mã đề {version}");

                        updateOverlay?.Invoke($"Chèn template (mã đề {version})", versionStart + stepPerVersion * 0.4);
                        await _openXMLService.InsertTemplateAsync(mixTemplate, mixDoc, mixInfo, version);
                        AddLog($"Chèn template xong cho mã đề {version}");

                        foreach (var a in answers)
                        {
                            a.Version = version;
                            allAnswers.Add(a);
                        }

                        updateOverlay?.Invoke($"Thêm footer & endnotes (mã đề {version})", versionStart + stepPerVersion * 0.6);
                        await _openXMLService.AddFooterAsync(mixDoc, version);
                        await _openXMLService.AddEndNotesAsync(mixDoc);
                        AddLog($"Thêm footer & endnotes xong cho mã đề {version}");

                        updateOverlay?.Invoke($"Định dạng lại văn bản đề (mã đề {version})", versionStart + stepPerVersion * 0.8);
                        await _openXMLService.FormatAllParagraphsAsync(mixDoc, mixInfo);
                        AddLog($"Định dạng văn bản đề xong cho mã đề {version}");

                        updateOverlay?.Invoke($"Xuất file đáp án (mã đề {version})", versionEnd);
                        await _openXMLService.AppendGuideAsync(answerDoc, answers, mixInfo, version);
                        await _openXMLService.MoveEssayTableToEndAsync(answerDoc);
                        await _openXMLService.FormatAllParagraphsAsync(answerDoc, mixInfo);
                        AddLog($"Xuất file đáp án xong cho mã đề {version}");

                        mixDoc.MainDocumentPart.Document.Save();
                        answerDoc.MainDocumentPart.Document.Save();
                        AddLog($"Lưu file đề và đáp án cho mã đề {version}");
                    }

                    await _interopWordService.UpdateFieldsAsync(mixFile);
                    AddLog($"Cập nhật số trang xong cho mã đề {version}");
                }

                updateOverlay?.Invoke("Xuất Excel đáp án", 95);
                _excelService.ExportExcelAnswers($"{outputFolder}\\{Constants.EXCEL_ANSWER_FILE}", allAnswers);
                AddLog("Xuất Excel đáp án xong");

                updateOverlay?.Invoke("Hoàn tất trộn đề", 100);
                AddLog("Trộn đề hoàn tất");
            });
        }

        private void ShowOverlay(string title, string statusText, double progressValue = 0)
        {
            ProgressOverlay.IsVisible = true;
            ProgressOverlay.Title = title;           // hiển thị tiêu đề ngắn gọn
            ProgressOverlay.StatusText = statusText; // hiển thị chi tiết tiến trình
            ProgressOverlay.ProgressValue = progressValue;

            // Nếu có giá trị tiến độ cụ thể thì không để indeterminate
            ProgressOverlay.IsIndeterminate = progressValue <= 0;
        }

        private async Task HideOverlayAsync(int delayMs = 1000)
        {
            await Task.Delay(delayMs);

            ProgressOverlay.IsVisible = false;
            ProgressOverlay.ProgressValue = 0;          // reset tiến trình
            ProgressOverlay.IsIndeterminate = true;     // đưa về trạng thái mặc định
            ProgressOverlay.StatusText = string.Empty;  // xóa nội dung chi tiết
            ProgressOverlay.Title = string.Empty;       // xóa tiêu đề
        }

        private void UpdateStatistics()
        {
            MultipleChoiceCount = Questions.Count(q => q.QuestionType == QuestionType.MultipleChoice);
            TrueFalseCount = Questions.Count(q => q.QuestionType == QuestionType.TrueFalse);
            ShortAnswerCount = Questions.Count(q => q.QuestionType == QuestionType.ShortAnswer);
            EssayCount = Questions.Count(q => q.QuestionType == QuestionType.Essay);
            UnknownCount = Questions.Count(q => q.QuestionType == QuestionType.Unknown);

            HasMultipleChoice = MultipleChoiceCount > 0;
            HasTrueFalse = TrueFalseCount > 0;
            HasShortAnswer = ShortAnswerCount > 0;
            HasEssay = EssayCount > 0;
            HasTotalPoint = HasMultipleChoice || HasTrueFalse || HasShortAnswer || HasEssay;

            float point = 0;
            if (MixInfo != null)
            {
                point += ParsePoint(MixInfo.PointMultipleChoice, HasMultipleChoice);
                point += ParsePoint(MixInfo.PointTrueFalse, HasTrueFalse);
                point += ParsePoint(MixInfo.PointShortAnswer, HasShortAnswer);
                point += ParsePoint(MixInfo.PointEssay, HasEssay);
            }
            TotalPoint = point > 0 ? point.ToString() : string.Empty;

            AddLog($"Thống kê: MC={MultipleChoiceCount}, TF={TrueFalseCount}, SA={ShortAnswerCount}, Essay={EssayCount}, Unknown={UnknownCount}");
        }

        private float ParsePoint(string pointValue, bool condition)
        {
            if (condition && !string.IsNullOrEmpty(pointValue))
            {
                string normalized = pointValue.Replace(',', '.');
                if (float.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out float result))
                {
                    return result;
                }
            }
            return 0;
        }
    }
}
