using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using Newtonsoft.Json.Linq;

namespace CKB
{
    internal static class KeepaAPI
    {
        private const string URI_ROOT = "https://api.keepa.com";
        private static readonly string SAVE_PATH_ROOT = $@"{EnvironmentSetup.DataRoot}\Keepa";
        private static readonly string RECORDS_DIR = $"{SAVE_PATH_ROOT}\\records";
        private static readonly string IMAGES_DIR = $"{SAVE_PATH_ROOT}\\images";
        private static readonly string API_REQUESTS_DIR = $@"{SAVE_PATH_ROOT}\\api_lookup_requests";
        private static readonly string API_RESPONSES_DIR = $@"{SAVE_PATH_ROOT}\\api_lookup_responses";
        private static readonly string API_BOOK_LASTLOOKUP_DIR = $@"{SAVE_PATH_ROOT}\\api_book_lastlookup";

        internal static IEnumerable<string> directories()
        {
            yield return SAVE_PATH_ROOT;
            yield return RECORDS_DIR;
            yield return IMAGES_DIR;
            yield return API_REQUESTS_DIR;
            yield return API_RESPONSES_DIR;
            yield return API_BOOK_LASTLOOKUP_DIR;
        }
        
        private const string LAST_CALL_DATEFORMAT = "yyyyMMdd_HHmmss";
        
        private static readonly Lazy<string> _apiKey = new Lazy<string>(() => Environment.GetEnvironmentVariable("KEEPA_API_KEY", EnvironmentVariableTarget.User));

        private static Lazy<HttpClient> _client = new Lazy<HttpClient>(() =>
        {
            var handler = new HttpClientHandler();
            if (handler.SupportsAutomaticDecompression)
            {
                handler.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
            }

            return  new HttpClient(handler);
        });
        
        private static Lazy<WebClient> _webClient = new Lazy<WebClient>(() => new WebClient());

        private static DateTime _lastRequestTime;
        private static void blockIfNecessary()
        {
            var nextRequestTime = _lastRequestTime.AddSeconds(2);

            var untilNextAvailableTime = (DateTime.Now - nextRequestTime);

            if (untilNextAvailableTime.TotalSeconds < 0)
            {
                Console.WriteLine($"Waiting {untilNextAvailableTime.TotalSeconds} seconds so don't hit api request limit.");
                Thread.Sleep(Math.Abs(untilNextAvailableTime.Milliseconds));
            }
        }

        private static string getResponse(string urlEnding_)
        {
            blockIfNecessary();
            var response = _client.Value.GetStringAsync($"{URI_ROOT}{urlEnding_}").Result;
            _lastRequestTime = DateTime.Now;
            return response;
        }

        public static IDictionary<string,KeepaRecord> GetDetailsForIdentifiers(string[] identifier_, bool forceRefresh_=false, bool ensureImagesToo_=true)
        {
            var toLookup = identifier_;

            var makeCall = true;
            
            if (!forceRefresh_)
            {
                var subset = identifier_.Where(x => !HaveLocalRecord(x));

                if (!subset.Any())
                    makeCall = false;
                else
                    toLookup = subset.ToArray();
            }

            if (makeCall)
            {
                var combinedIdentifiers = string.Join(",", toLookup);
                var url = $@"/product?key={_apiKey.Value}&domain=2&code={combinedIdentifiers}";

                var time = DateTime.Now;

                File.WriteAllText($@"{API_REQUESTS_DIR}\lookup-{time:yyyyMMdd-HHmmss}.json", combinedIdentifiers);

                var response = getResponse(url);

                if (response != null)
                {
                    toLookup.ForEach(id =>
                        File.WriteAllText(lastLookupPath(id), time.ToString(LAST_CALL_DATEFORMAT)));

                    writeRecordsFromResponse(response, ensureImagesToo_);
                } 

                File.WriteAllText($@"{API_RESPONSES_DIR}\lookup-{time:yyyyMMdd-HHmmss}.json", response);
            }

            return identifier_.ToDictionary(x => x, x => GetRecordForIdentifier(x));
        }

        private static string recordPath(string identifier_) => $@"{RECORDS_DIR}\{identifier_}.json";
        private static string lastLookupPath(string identifier_)=>$@"{API_BOOK_LASTLOOKUP_DIR}\{identifier_}.txt";
        
        public static bool HaveLocalRecord(string identifier_) => File.Exists(recordPath(identifier_));

        public static void ReWritefromResponses()
        {
            var files = Directory.GetFiles(API_RESPONSES_DIR, "*.json");
            
            files.ForEach(x=>writeRecordsFromResponse(File.ReadAllText(x),true));
        }

        private static void writeRecordsFromResponse(string response_, bool ensureImagesToo_)
        {
            var doc = JObject.Parse(response_);

            var productToken = doc["products"];

            if (productToken == null) return;

            var productArray = (JArray)productToken;

            productArray.ForEach(product =>
            {
                writeProductRecord(product);
                
                if (ensureImagesToo_)
                    KeepaRecord.Parse(product).DownloadImages();
            });
        }

        private static KeepaRecord writeProductRecord(JToken product)
        {
            var p = KeepaRecord.Parse(product);

            if (!string.IsNullOrEmpty(p.Asin))
            {
                $"Writing keepa record for asin identifier to '{recordPath(p.Asin)}'".ConsoleWriteLine();
                File.WriteAllText(recordPath(p.Asin), product.ToString());
            }

            p.EanList?.ForEach(x=>
            {
                $"Writing record for ean identifier to '{recordPath(x)}'".ConsoleWriteLine();
                File.WriteAllText(recordPath(x), product.ToString());
            });

            return p;
        }

        public static KeepaRecord GetRecordForIdentifier(string identifier_)
        {
            var path = recordPath(identifier_);

            return File.Exists(recordPath(identifier_)) ? KeepaRecord.Parse( JToken.Parse(File.ReadAllText(path))) : null;
        }

        public static IEnumerable<KeepaRecord> AllLocalRecords()
            => Directory.GetFiles(RECORDS_DIR, "*.json").Select(x => KeepaRecord.Parse(JToken.Parse(File.ReadAllText(x))));

        public static DateTime? LastLookupTime(string isbn_)
        {
            var path = lastLookupPath(isbn_);

            if (!File.Exists(path))
                return default;

            return DateTime.ParseExact(File.ReadAllText(path), LAST_CALL_DATEFORMAT, CultureInfo.InvariantCulture);
        }
        public static string[] ImagePaths(this KeepaRecord record_)
            => record_.ImagesCSV.Split(',').Select(part => $@"{IMAGES_DIR}\{part}").ToArray();

        public static void DownloadImages(this KeepaRecord record_, bool force_ = false)
        {
            if (string.IsNullOrEmpty(record_?.ImagesCSV))
                return;
            
            var paths = record_.ImagePaths();

            paths.ForEach(path =>
            {
                if (File.Exists(path) && !force_)
                    return;

                var amazonPath = $@"https://images-na.ssl-images-amazon.com/images/I/{record_.ImagesCSV.Split(',')[0]}";

                try
                {
                    _webClient.Value.DownloadFile(amazonPath, path);
                    
                    if(File.Exists(path))
                        $"Wrote image file to '{path}'".ConsoleWriteLine();
                }
                catch (Exception ex_)
                {
                    Console.WriteLine(ex_.Message);
                }
            });
        }
    }

    public class KeepaRecord
    {
        public static KeepaRecord Parse(JToken token_)
        {
            var ret = new KeepaRecord
            {
                Asin = token_.ExtractString("asin"),
                Title = token_.ExtractString("title"),
                ImagesCSV = token_.ExtractString("imagesCSV"),
                Type=token_.ExtractString("type"),
                Binding = token_.ExtractString("binding"),
                Description = token_.ExtractString("description"),
                Manufacturer = token_.ExtractString("manufacturer"),
                Author = token_.ExtractString("author")
            };

            try
            {
                var eanToken = (JArray)token_["eanList"];
                ret.EanList = eanToken.Select(x => x.ToString()).ToArray();
            }
            catch
            {
            }

            try
            {
                var category = (JArray) token_["categoryTree"];
                ret.CategoryTree = category.Select(x => x.ExtractString("name")).ToArray();
            }
            catch 
            {
            }

            return ret;
        }
        
        /*
         * Usage and structure. Each product sold on Amazon.com is given a unique ASIN.
         * For books with 10-digit International Standard Book Number (ISBN), the ASIN
         * and the ISBN are the same. The Kindle edition of a book will not use its ISBN
         * as the ASIN, although the electronic version of a book may have its own ISBN.
         */
        
        
        
        public string Asin { get; set; }
        public string Title { get; set; }
        public string ImagesCSV { get; set; }
        public string Author { get; set; }
        public string Type { get; set; }
        public string Binding { get; set; }
        public string Description { get; set; }
        public string Manufacturer { get; set; }

        // https://networlding.com/ean-vs-isbn-what-you-need-to-know/#:~:text=These%20two%20numbers%20are%20essentially,a%20sticker%20on%20the%20book.
        public string[] EanList { get; set; }
        
        public string[] CategoryTree { get; set; }
    }

}
