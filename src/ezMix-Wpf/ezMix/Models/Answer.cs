namespace ezMix.Models
{
    public class Answer
    {
        public string AnswerText { get; set; } = string.Empty;
        public bool IsCorrect { get; set; }
        public string FilePath { get; set; } = string.Empty;
    }
}
