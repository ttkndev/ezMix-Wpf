using DocumentFormat.OpenXml;
using ezMix.Models.Enums;
using System.Collections.Generic;

namespace ezMix.Models
{
    public class QuestionExport
    {
        public string Version { get; set; }
        public int QuestionNumber { get; set; }
        public string CorrectAnswer { get; set; }
        public QuestionType Type { get; set; }
        public string Point { get; set; }
        public List<OpenXmlElement> AnswerElements { get; set; }
    }
}
