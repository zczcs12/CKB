using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace CKB
{
    public static class AbeBooksAPI
    {
        private static readonly string SAVE_PATH_ROOT = $@"{EnvironmentSetup.DataRoot}\AbeBooks\Images";

        private static Lazy<WebClient> _webClient = new Lazy<WebClient>(() => new WebClient());

        internal static IEnumerable<string> directories()
        {
            yield return SAVE_PATH_ROOT;
        }
        
        private static string localPathToImage(string identifier_) => $"{SAVE_PATH_ROOT}\\{identifier_}.jpg";

        public static string TryGetImageForIdentifier(string identifier_)
        {
            var path = localPathToImage(identifier_);

            if (File.Exists(path))
                return path;

            try
            {
                _webClient.Value.DownloadFile($"https://pictures.abebooks.com/isbn/{identifier_}-uk.jpg",path);
                $"Wrote image file for {identifier_} to '{path}'".ConsoleWriteLine();
            }
            catch 
            {
            }

            return File.Exists(path) ? path : null;
        }

    }
}