using ezMix.Services.Interfaces;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;

namespace ezMix.Services.Implementations
{
    public class GeminiService : IGeminiService
    {
        private readonly HttpClient _httpClient;
        private readonly string endpoint = "https://generativelanguage.googleapis.com/v1/models/{0}:generateContent?key={1}";

        public GeminiService()
        {
            _httpClient = new HttpClient()
            {
                Timeout = TimeSpan.FromMinutes(2)
            };
        }

        //Gọi mô hình Gemini với một đoạn prompt(văn bản đầu vào).
        public async Task<string> CallGeminiAsync(string model, string apiKey, string prompt)
        {
            var url = string.Format(endpoint, model, apiKey);

            // Body JSON đúng chuẩn
            var requestBody = new
            {
                contents = new[]
                {
                    new
                    {
                        parts = new[]
                        {
                            new { text = prompt }
                        }
                    }
                }
            };

            var json = JsonSerializer.Serialize(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            try
            {
                var response = await _httpClient.PostAsync(url, content);
                var responseString = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    // Trả về chi tiết lỗi từ Google
                    return $"❌ API trả về lỗi {response.StatusCode}: {responseString}";
                }

                // Parse JSON để lấy text
                using (var doc = JsonDocument.Parse(responseString))
                {
                    var root = doc.RootElement;

                    if (root.TryGetProperty("candidates", out var candidates) &&
                        candidates.GetArrayLength() > 0 &&
                        candidates[0].TryGetProperty("content", out var contentProp) &&
                        contentProp.TryGetProperty("parts", out var parts) &&
                        parts.GetArrayLength() > 0 &&
                        parts[0].TryGetProperty("text", out var textElement))
                    {
                        return textElement.GetString() ?? string.Empty;
                    }
                }

                return "❌ Không tìm thấy nội dung trả về từ Gemini.";
            }
            catch (HttpRequestException ex)
            {
                return $"❌ Lỗi kết nối API: {ex.Message}";
            }
            catch (JsonException ex)
            {
                return $"❌ Lỗi parse JSON: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"❌ Lỗi không xác định: {ex.Message}";
            }
        }

        //Gọi mô hình Gemini để trích xuất văn bản từ một hình ảnh.
        public async Task<string> CallGeminiExtractTextFromImageAsync(string model, string apiKey, string imagePath, string prompt)
        {
            if (string.IsNullOrEmpty(prompt))
                prompt = "Hãy nhận diện văn bản trong ảnh này và xuất ra Markdown";

            using (var client = new HttpClient())
            {
                var imageBytes = System.IO.File.ReadAllBytes(imagePath);
                string base64Image = Convert.ToBase64String(imageBytes);

                var requestBody = new
                {
                    contents = new[]
                    {
                        new {
                            parts = new object[]
                            {
                                new { text = prompt },
                                new { inline_data = new { mime_type = "image/png", data = base64Image } }
                            }
                        }
                    }
                };

                string json = JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var url = string.Format(endpoint, model, apiKey);
                var response = await client.PostAsync(url, content);
                string result = await response.Content.ReadAsStringAsync();

                return result;
            }
        }

        // Gọi mô hình Gemini để trích xuất văn bản từ một file PDF.
        public async Task<string> CallGeminiExtractTextFromPdfAsync(string model, string apiKey, string pdfPath, string prompt)
        {
            if (string.IsNullOrEmpty(prompt))
                prompt = "Hãy nhận diện văn bản trong ảnh này và xuất ra Markdown";

            using (var client = new HttpClient())
            {
                var pdfBytes = System.IO.File.ReadAllBytes(pdfPath);
                string base64Pdf = Convert.ToBase64String(pdfBytes);

                var requestBody = new
                {
                    contents = new[]
                    {
                new {
                    parts = new object[]
                    {
                        new { text = prompt },
                        new { inline_data = new { mime_type = "application/pdf", data = base64Pdf } }
                    }
                }
            }
                };

                string json = JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // Endpoint phải kèm API key
                var url = string.Format(endpoint, model, apiKey);
                var response = await client.PostAsync(url, content);
                string result = await response.Content.ReadAsStringAsync();

                return result;
            }
        }
    }
}
