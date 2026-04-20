using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace ezMix.Helpers
{
    public class InternetHelper
    {
        public static Task<bool> IsInternetAvailableAsync()
        {
            return Task.Run(() =>
            {
                try
                {
                    var request = (HttpWebRequest)WebRequest.Create("http://www.google.com");
                    request.Timeout = 3000; // 3 giây
                    request.Method = "GET";

                    using (var response = (HttpWebResponse)request.GetResponse())
                    {
                        return response.StatusCode == HttpStatusCode.OK;
                    }
                }
                catch
                {
                    return false;
                }
            });
        }
    }
}
