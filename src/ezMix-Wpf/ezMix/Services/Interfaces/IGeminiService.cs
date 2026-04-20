using System.Threading.Tasks;

namespace ezMix.Services.Interfaces
{
    public interface IGeminiService
    {
        Task<string> CallGeminiAsync(string model, string apiKey, string prompt);
        Task<string> CallGeminiExtractTextFromImageAsync(string model, string apiKey, string imagePath, string prompt);
        Task<string> CallGeminiExtractTextFromPdfAsync(string model, string apiKey, string pdfPath, string prompt);
    }
}
