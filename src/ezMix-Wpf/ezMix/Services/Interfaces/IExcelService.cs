using ezMix.Models;
using System.Collections.Generic;

namespace ezMix.Services.Interfaces
{
    public interface IExcelService
    {
        void ExportExcelAnswers(string filePath, List<QuestionExport> answers);
    }
}
