using ezMix.Models;
using System.Threading.Tasks;

namespace ezMix.Services.Interfaces
{
    public interface IUpdateCheckService
    {
        Task<UpdateInfo> GetLatestAsync(string url);

        bool HasUpdate(string currentVersion, string latestVersion);
    }
}
