using ezMix.Models;
using ezMix.Services.Interfaces;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace ezMix.Services.Implementations
{
    public class UpdateCheckService : IUpdateCheckService
    {
        private readonly HttpClient _http = new HttpClient();

        public async Task<UpdateInfo> GetLatestAsync(string url)
        {
            var response = await _http.GetAsync(url);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<UpdateInfo>(json);
        }

        public bool HasUpdate(string current, string latest)
        {
            return new Version(latest) > new Version(current);
        }
    }
}
