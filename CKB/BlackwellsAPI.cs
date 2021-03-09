using System;
using System.IO;
using System.Net;

namespace CKB
{
    public static class BlackwellsAPI
    {
        // https://blackwells.co.uk/jacket/l/9781782096580.jpg"
        
        private static readonly string SAVE_PATH_ROOT = $@"{Environment.GetEnvironmentVariable("CKB_DATA_ROOT",EnvironmentVariableTarget.User)}\Blackwells\Images";

        private static Lazy<WebClient> _webClient = new Lazy<WebClient>(() => new WebClient());

        private static string localPathToImage(string identifier_) => $"{SAVE_PATH_ROOT}\\{identifier_}.jpg";

        public static string TryGetImageForIdentifier(string identifier_)
        {
            var path = localPathToImage(identifier_);

            if (File.Exists(path))
                return path;

            try
            {
                _webClient.Value.DownloadFile($"https://blackwells.co.uk/jacket/l/{identifier_}.jpg",path);
                $"Wrote image file for {identifier_} to '{path}'".ConsoleWriteLine();
            }
            catch 
            {
            }

            return File.Exists(path) ? path : null;
        }

    }
}