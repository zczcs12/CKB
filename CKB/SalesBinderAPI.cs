using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using CsvHelper.Configuration.Attributes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CKB
{
    internal static class SalesBinderAPI
    {

        private const string URI_ROOT = "https://ckb.salesbinder.com/api/2.0/";
        internal static readonly string SAVE_PATH_ROOT = $@"{EnvironmentSetup.DataRoot}\SalesBinder";


        internal static IEnumerable<string> directories()
        {
            yield return SAVE_PATH_ROOT;
            yield return ImageDirectory;
            yield return InventoryItemsDirectory;
        }

        private static string InventoryFilePath => $@"{SAVE_PATH_ROOT}\inventory.json";
        private static string ContactsFilePath = $@"{SAVE_PATH_ROOT}\contacts.json";
        private static string AccountsFilePath = $@"{SAVE_PATH_ROOT}\accounts.json";
        private static string InvoicesFilePath = $@"{SAVE_PATH_ROOT}\invoices.json";
        private static string SettingsFilePath = $@"{SAVE_PATH_ROOT}\settings.json";
        private static string LocationsFilePath = $@"{SAVE_PATH_ROOT}\locations.json";

        private static string InventoryItemsDirectory = $@"{SAVE_PATH_ROOT}\inventoryrecords";
        private static string ImageDirectory = $@"{SAVE_PATH_ROOT}\images";

        #region Inventory

        
        private static string InventoryInvidualItemPath(string item_id_) =>
            $@"{InventoryItemsDirectory}\{item_id_}.json";

        private static readonly Lazy<SalesBinderInventoryItem[]> _inventory = new Lazy<SalesBinderInventoryItem[]>(() =>
            File.Exists(InventoryFilePath)
                ? JsonConvert.DeserializeObject<SalesBinderInventoryItem[]>(File.ReadAllText(InventoryFilePath))
                : null);

        private static readonly Lazy<Dictionary<string, SalesBinderInventoryItem>> _inventoryById =
            new Lazy<Dictionary<string, SalesBinderInventoryItem>>(
                () => _inventory.Value.ToDictionary(x => x.Id, x => x)
            );

        private static readonly Lazy<Dictionary<string, SalesBinderInventoryItem>> _inventoryByBarcode =
            new Lazy<Dictionary<string, SalesBinderInventoryItem>>(
                () => _inventory.Value.GroupBy(x => x.BarCode)
                    .ToDictionary(x => x.Key, x => x.OrderByDescending(b => b.Quantity).First())
            );

        public static SalesBinderInventoryItem[] Inventory => _inventory.Value;
        public static IDictionary<string, SalesBinderInventoryItem> InventoryById => _inventoryById.Value;
        public static IDictionary<string, SalesBinderInventoryItem> InventoryByBarcode => _inventoryByBarcode.Value;



        #endregion

        #region Contacts

        private static Lazy<SalesBinderContact[]> _contacts = new Lazy<SalesBinderContact[]>(() =>
            File.Exists(ContactsFilePath)
                ? JsonConvert.DeserializeObject<SalesBinderContact[]>(File.ReadAllText(ContactsFilePath))
                : null);

        public static SalesBinderContact[] Contacts => _contacts.Value;

        #endregion

        #region Accounts

        private static readonly Lazy<SalesBinderAccount[]> _accounts = new Lazy<SalesBinderAccount[]>(() =>
            File.Exists(AccountsFilePath)
                ? JsonConvert.DeserializeObject<SalesBinderAccount[]>(File.ReadAllText(AccountsFilePath))
                : null);

        private static readonly Lazy<Dictionary<string, SalesBinderAccount>> _accountsById =
            new Lazy<Dictionary<string, SalesBinderAccount>>(
                () => _accounts.Value.ToDictionary(x => x.Id, x => x)
            );

        public static SalesBinderAccount[] Accounts => _accounts.Value;
        public static IDictionary<string, SalesBinderAccount> AccountById => _accountsById.Value;

        #endregion

        #region Invoices

        private static readonly Lazy<SalesBinderInvoice[]> _invoices = new Lazy<SalesBinderInvoice[]>(() =>
            File.Exists(InvoicesFilePath)
                ? JsonConvert.DeserializeObject<SalesBinderInvoice[]>(File.ReadAllText(InvoicesFilePath))
                : null);

        public static SalesBinderInvoice[] Invoices => _invoices.Value;

        #endregion

        #region Units of Measure

        private static readonly Lazy<SalesBinderUnitOfMeasure[]> _unitsOfMeasure = new Lazy<SalesBinderUnitOfMeasure[]>(
            () =>
            {
                var s = JToken.Parse(ExtensionMethods.GetTextFromEmbeddedResource("CKB.Static.units_of_measure.json"));
                var arr = s?.ExtractArray("data");
                return arr == null ? null : arr.Select(SalesBinderUnitOfMeasure.Parse).ToArray();
            });

        public static SalesBinderUnitOfMeasure[] UnitsOfMeasure => _unitsOfMeasure.Value;

        private static readonly Lazy<Dictionary<string, SalesBinderUnitOfMeasure>> _unitsOfMeasureByName =
            new Lazy<Dictionary<string, SalesBinderUnitOfMeasure>>(
                () => UnitsOfMeasure.ToDictionary(x => x.ShortName, x => x)
            );

        public static IDictionary<string, SalesBinderUnitOfMeasure> UnitsOfMeasureByName => _unitsOfMeasureByName.Value;

        #endregion

        private static readonly Lazy<Dictionary<string, string>> _customFieldCodes =
            new Lazy<Dictionary<string, string>>(
                () =>
                {
                    var s = JToken.Parse(ExtensionMethods.GetTextFromEmbeddedResource("CKB.Static.custom_fields.json"));
                    var arr = s?.ExtractArray("data");

                    return arr.ToDictionary(x => x.ExtractString("name"), x => x.ExtractString("custom_field_id"));
                }
            );

        public static Dictionary<string, string> CustomFieldNameToId => _customFieldCodes.Value;
        
        private static readonly Lazy<HttpClient> _client = new Lazy<HttpClient>(() =>
        {
            var apiKey = Environment.GetEnvironmentVariable("SALESBINDER_API_KEY", EnvironmentVariableTarget.User);

            var client = new HttpClient();
            var byteArray = Encoding.ASCII.GetBytes($"{apiKey}:x");
            client.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            return client;
        });

        private static Lazy<WebClient> _webClient = new Lazy<WebClient>(() => new WebClient());


        public static IEnumerable<SalesBinderInventoryItem> RetrieveAndSaveInventory(bool topup_)
            => retrieveAndSave(() =>
            {
                var current = Inventory;

                if (!topup_ || !current.Any()) return RetrieveInventory();

                var lastUpdate = new FileInfo(InventoryFilePath).LastWriteTimeUtc;

                var topup = RetrieveInventory(lastUpdate.AddDays(-1d));

                if (!topup.Any()) return current;

                var currentD = current.ToDictionary(x => x.Id, x => x);
                
                topup.ForEach(x => currentD[x.Id] = x);

                return currentD.Values;

            }, InventoryFilePath);

        public static IEnumerable<SalesBinderInventoryItem> RetrieveInventory(DateTime? since_ = null)
            => since_.HasValue
                ? retrievePaginated("items", "items", SalesBinderInventoryItem.Parse,
                    $"&pageLimit=100&modifiedSince={since_.Value.ToEpochTime()}", false)
                : retrievePaginated("items", "items", SalesBinderInventoryItem.Parse, "&pageLimit=100&", true);

        private static IEnumerable<T> retrievePaginated<T>(string apiName_, string groupTokenName_,
            Func<JToken, T> creator_, string optionalArgs_ = null, bool log_=true)
        {
            var allItems = new List<T>();

            var pageNumber = 1;
            var stop = false;

            while (!stop)
            {
                var response = getResponse($"{apiName_}.json?page={pageNumber}{optionalArgs_}");

                var j = JObject.Parse(response);
                var pagesToken = j["pages"];
                var items = (JArray) j[groupTokenName_].First;

                allItems.AddRange(items.Select(x => creator_(x)));

                if (pagesToken == null)
                    stop = true;
                else
                {
                    var totalPages = int.Parse(pagesToken.ToString());

                    if(log_)
                        Console.WriteLine($"Processed page {pageNumber} of {totalPages} of {apiName_}");

                    if (pageNumber >= totalPages)
                        stop = true;
                    else
                        pageNumber += 1;
                }
            }

            return allItems;

        }

        public static IEnumerable<SalesBinderContact> RetrieveContacts()
            => retrievePaginated("contacts", "contacts", x => SalesBinderContact.Parse(x));

        public static void RetrieveAndSaveContacts()
            => retrieveAndSave(RetrieveContacts, ContactsFilePath);

        public static IEnumerable<SalesBinderAccount> RetrieveAccounts()
            => retrievePaginated("customers", "customers", x => SalesBinderAccount.Parse(x), "&pageLimit=200");

        public static void RetrieveAndSaveAccounts()
            => retrieveAndSave(RetrieveAccounts, AccountsFilePath);

        public static IEnumerable<SalesBinderInvoice> RetrieveInvoices(DateTime? since_ = null)
            => since_.HasValue
                ? retrievePaginated("documents", "documents", x => SalesBinderInvoice.Parse(x),
                    $"&contextId=5&pageLimit=200&modifiedSince={since_.Value.ToEpochTime()}")
                : retrievePaginated("documents", "documents", x => SalesBinderInvoice.Parse(x),
                    "&contextId=5&pageLimit=200");

        public static void CreateInventory(string barCode_, KeepaRecord useTheseDetails_)
        {
            var toAdd = new JObject(
                new JProperty("item", new JObject(
                    new JProperty(InventoryFields.Name, useTheseDetails_.Title),
                    new JProperty("description", useTheseDetails_.Description),
                    new JProperty(InventoryFields.SKU, barCode_),
                    new JProperty(InventoryFields.BarCode, barCode_),
                    new JProperty("multiple", 1),
                    new JProperty("threshold", 0),
                    new JProperty(InventoryFields.Cost, 0.0),
                    new JProperty(InventoryFields.Price, 0.0),
                    new JProperty(InventoryFields.Quantity, 0),
                    new JProperty("category_id", "5b1aafe1-ed18-43b2-b2dd-2ed40a8e0005"),
                    new JProperty("item_details", new JArray
                    {
                        new JObject(
                            new JProperty("custom_field_id", InventoryCustomFields.Author),
                            new JProperty("value", useTheseDetails_.Author)
                        ),
                        new JObject(
                            new JProperty("custom_field_id", InventoryCustomFields.Publisher),
                            new JProperty("value", useTheseDetails_.Manufacturer)
                        ),
                    })
                )));

            var response = postJson("items.json", toAdd.ToString());
            
            $"{JObject.Parse(response)}".ConsoleWriteLine();
        }
        
        public static void RetrieveAndSaveInvoices(bool topup_)
            => retrieveAndSave(() =>
            {
                var current = Invoices;

                if (!topup_ || !current.Any()) return RetrieveInvoices();

                var lastUpdate = new FileInfo(InvoicesFilePath).LastWriteTimeUtc;

                var topup = RetrieveInvoices(lastUpdate.AddDays(-1d));

                if (!topup.Any()) return current;

                var currentD = current.ToDictionary(x => x.Id, x => x);
                topup.ForEach(x => currentD[x.Id] = x);

                return currentD.Values;

            }, InvoicesFilePath);

        private static IEnumerable<T> retrieveAndSave<T>(Func<IEnumerable<T>> func_, string filePath_)
        {
            var allItems = func_().ToArray();
            var txt = JsonConvert.SerializeObject(allItems.ToArray(), Formatting.None);
            File.WriteAllText(filePath_, txt);
            return allItems;
        }

        public static void SendItemUpdate(string itemId_, string json_)
        {
            var response = putJson($"items/{itemId_}.json", json_);

            var p = JObject.Parse(response);
            
            p.ExtractString("message").ConsoleWriteLine();
        }
        public static void RetrieveAndSaveSettings()
            => File.WriteAllText(SettingsFilePath, getResponse("settings.json"));

        public static void RetrieveAndSaveLocations()
            => File.WriteAllText(LocationsFilePath, getResponse("locations.json"));

        public static void DownloadBookImages()
        {
            Inventory
                .Where(x => x.Quantity > 0)
                .ForEach(item =>
                {
                    var sets = new (SalesBinderInventoryItem Book, string Url, string File)[]
                        {
                            (item, item.ImageURLSmall, item.ImageFilePath(ImageSize.Small)),
                            (item, item.ImageURLMedium, item.ImageFilePath(ImageSize.Medium)),
                            (item, item.ImageURLLarge, item.ImageFilePath(ImageSize.Large)),
                        }
                        .Where(x => !string.IsNullOrEmpty(x.Url) && !File.Exists(x.File));

                    if (!sets.Any()) return;

                    sets.ForEach(set =>
                    {
                        try
                        {
                            _webClient.Value.DownloadFile(set.Url, set.File);
                            $"Downloaded {set.File} for inventory item {set.Book.Name}".ConsoleWriteLine();
                        }
                        catch
                        {
                            $"Error downloading file from {set.Url} for book '{set.Book.Name}'".ConsoleWriteLine();
                        }
                    });
                });
        }

        private static DateTime _lastRequestTime;

        private static void blockIfNecessary()
        {
            /* salesbinder support:
             * Our current rate limits are as follows:
                - 18 requests per 10 seconds, Block for 1 minute
                - 60 requests per 1 minute, Block for 1 minute
             */

            Thread.Sleep(2000);

            // var nextRequestTime = _lastRequestTime.AddSeconds(2);
            //
            // var untilNextAvailableTime = (DateTime.Now - nextRequestTime);
            //
            // if (untilNextAvailableTime.TotalSeconds < 0)
            // {
            //     Console.WriteLine($"Waiting {untilNextAvailableTime.TotalSeconds} seconds so don't hit api request limit.");
            //     Thread.Sleep(Math.Abs(untilNextAvailableTime.Milliseconds));
            // }
        }

        private static string getResponse(string urlEnding_)
        {
            blockIfNecessary();
            var response = _client.Value.GetStringAsync($"{URI_ROOT}{urlEnding_}").Result;
            _lastRequestTime = DateTime.Now;
            return response;
        }

        private static string putJson(string urlEnding_, string json_)
        {
            blockIfNecessary();
            var httpContent = new StringContent(json_, Encoding.UTF8, "application/json");
            var uri = new Uri($"{URI_ROOT}{urlEnding_}");
            var response = _client.Value.PutAsync(uri, httpContent).Result;
            _lastRequestTime = DateTime.Now;
            return response?.Content.ReadAsStringAsync().Result;
        }

        private static string postJson(string urlEnding_, string json_)
        {
            var httpContent = new StringContent(json_, Encoding.UTF8, "application/json");
            var uri = new Uri($"{URI_ROOT}{urlEnding_}");
            var response = _client.Value.PostAsync(uri, httpContent).Result;
            _lastRequestTime = DateTime.Now;
            return response?.Content.ReadAsStringAsync().Result;
        }

        private static string postImage(string urlEnding_, string filePath_)
        {
            blockIfNecessary();
            var uri = new Uri($"{URI_ROOT}{urlEnding_}");

            var request = new HttpRequestMessage(HttpMethod.Post, uri);
            request.Content = new ByteArrayContent(File.ReadAllBytes(filePath_));
            request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("image/jpeg");

            var response = _client.Value.SendAsync(request).Result;
            _lastRequestTime = DateTime.Now;

            return response?.Content.ReadAsStringAsync().Result;
        }

        public static JObject RetrieveJsonForInventoryItem(string itemId_)
        {
            var response = getResponse($"items/{itemId_}.json");

            return JObject.Parse(response);
        }

        public static void UploadImage(string identifier_, string imagePath_)
        {
            if (!_inventoryByBarcode.Value.TryGetValue(identifier_, out var record))
            {
                $"Don't know of an inventory item with barcode {identifier_} so can't update it."
                    .ConsoleWriteLine();
                return;
            }

            if (!File.Exists(imagePath_))
            {
                $"The image path '{imagePath_}' does not exist"
                    .ConsoleWriteLine();
                return;
            }

            var urlEnding = $"images/upload/{record.Id}.json";

            var response = postImage(urlEnding, imagePath_);

            response.ConsoleWriteLine();
        }
    }

    public enum ImageSize
    {
        Small,
        Medium,
        Large
    }

    public class SalesBinderInventoryItem
    {
        public static SalesBinderInventoryItem Parse(JToken token_)
        {
            var ret = new SalesBinderInventoryItem
            {
                Name = token_.ExtractString(InventoryFields.Name),
                Quantity = token_.ExtractLong(InventoryFields.Quantity),
                Cost = token_.ExtractDecimal(InventoryFields.Cost),
                Price = token_.ExtractDecimal(InventoryFields.Price),
                SKU = token_.ExtractString(InventoryFields.SKU),
                ItemNumber = token_.ExtractLong(InventoryFields.ItemNumber),
                BarCode = token_.ExtractString(InventoryFields.BarCode),
                Id = token_.ExtractString(InventoryFields.Id),
            };

            try
            {
                var unitOfMeasure = token_["unit_of_measure"];
                ret.BinLocation = unitOfMeasure?.ExtractString("full_name");
                ret.BinLocationId = unitOfMeasure?.ExtractString("id");
            }
            catch
            {
            }

            try
            {
                if(token_["item_details"] is JArray details)
                    foreach (var det in details)
                    {
                        var field = det["custom_field"]["name"].ToString();
                        var val = det["value"].ToString();

                        if (_customFields.TryGetValue(field, out var f))
                            f?.Invoke(ret, val);
                    }
            }
            catch
            {
            }

            var images = token_["images"];

            ret.ImageURLSmall = images?.First.ExtractString("url_small");
            ret.ImageURLMedium = images?.First.ExtractString("url_medium");
            ret.ImageURLLarge = images?.First.ExtractString("url_large");

            return ret;
        }

        private static readonly Dictionary<string, Action<SalesBinderInventoryItem, string>> _customFields =
            new Dictionary<string, Action<SalesBinderInventoryItem, string>>
            {
                {InventoryCustomFields.ProductType, (i, v) => i.ProductType = v},
                {InventoryCustomFields.ProductType2, (i, v) => i.ProductType2 = v},
                {InventoryCustomFields.ProductType3, (i, v) => i.ProductType3 = v},
                {InventoryCustomFields.Style, (i, v) => i.Style = v},
                {InventoryCustomFields.CountryOfOrigin, (i, v) => i.CountryOfOrigin = v},
                {InventoryCustomFields.Publisher, (i, v) => i.Publisher = v},
                {InventoryCustomFields.KidsAdult, (i, v) => i.KidsOrAdult = v},
                {InventoryCustomFields.Condition, (i, v) => i.Condition = v},
                {InventoryCustomFields.Author, (i, v) => i.Author = v},
                {InventoryCustomFields.PackSize, (i, v) => i.PackSize = v},
                {InventoryCustomFields.VAT, (i, v) => i.VAT = v},
                {InventoryCustomFields.FullRrp, (i, v) => i.FullRRP = v},
                {InventoryCustomFields.CommodityCode, (i, v) => i.CommodityCode = v},
                {InventoryCustomFields.MaterialComposition, (i, v) => i.MaterialComposition = v},
                {InventoryCustomFields.Clearance, (i, v) => i.Clearance = v},
                {InventoryCustomFields.BinLocation, null},
                {InventoryCustomFields.SalesRestrictions,(i,v)=>i.SalesRestrictions=v},
                {InventoryCustomFields.Rating, (i, v) => i.Rating = double.TryParse(v,out var val)? val : default(double?)}
            };


        [Index(0)]
        public string Name { get; set; }
        [Index(1)]
        public string SKU { get; set; }
        [Index(2)]
        public string BarCode { get; set; }
        [Index(3)]
        public long Quantity { get; set; }
        [Index(4)]
        public decimal Cost { get; set; }
        [Index(5)]
        public decimal Price { get; set; }
        [Index(6)]
        public string FullRRP { get; set; }
        [Index(7)]
        public string Author { get; set; }
        [Index(8)]
        public string Publisher { get; set; }
        [Ignore]
        public long ItemNumber { get; set; }
        [Index(9)]
        public string Style { get; set; }
        [Index(10)]
        public string CountryOfOrigin { get; set; }
        [Index(11)]
        public string KidsOrAdult { get; set; }
        [Index(12)]
        public string ProductType { get; set; }
        [Index(13)]
        public string ProductType2 { get; set; }
        [Index(14)]
        public string ProductType3 { get; set; }
        [Index(15)]
        public string Condition { get; set; }
        [Index(16)]
        public string Clearance { get; set; }
        [Index(17)]
        public string BinLocation { get; set; }
        [Ignore]
        public string BinLocationId { get; set; }
        [Index(18)]
        public string PackSize { get; set; }
        [Index(19)]
        public string VAT { get; set; }
        [Index(20)]
        public string CommodityCode { get; set; }
        [Index(21)]
        public string MaterialComposition { get; set; }
        [Index(22)]
        public double? Rating { get; set; }
        [Index(23)]
        public string SalesRestrictions { get; set; }
        [Index(24)]
        public string Id { get; set; }
        [Ignore]
        public string ImageURLSmall { get; set; }
        [Ignore]
        public string ImageURLMedium { get; set; }
        [Ignore]
        public string ImageURLLarge { get; set; }
    }

    public static class InventoryCustomFields
    {
        public const string ProductType = "Product Type";
        public const string ProductType2 = "Product Type 2";
        public const string ProductType3 = "Product Type 3";
        public const string Style = "Style";
        public const string CountryOfOrigin = "Country of Origin";
        public const string Publisher = "Publisher";
        public const string KidsAdult = "Kids/adult";
        public const string Condition = "Condition";
        public const string Author = "Author";
        public const string PackSize = "Pack size";
        public const string VAT = "VAT";
        public const string FullRrp = "Full RRP";
        public const string CommodityCode = "Commodity Code";
        public const string MaterialComposition = "Material Composition";
        public const string Clearance = "Clearance";
        public const string BinLocation = "Bin Location";
        public const string Rating = "Rating";
        public const string SalesRestrictions = "Sales Restrictions";
    }

    public static class InventoryFields
    {
        public const string Name = "name";
        public const string SKU = "sku";
        public const string BarCode = "barcode";
        public const string Quantity = "quantity";
        public const string Cost = "cost";
        public const string Price = "price";
        public const string ItemNumber = "item_number";
        public const string Id = "id";
    }

    public static class SalesBinderInventoryItemExtensions
    {
        private static void setItemDetail(JObject root, string customFieldname_, object value_)
        {
            var itemToken = root["item"] as JObject;

            var arr = itemToken["item_details"] as JArray;

            if (arr == null)
            {
                arr = new JArray();
                itemToken.Add("item_details",arr);
            }

            var customFieldId = SalesBinderAPI.CustomFieldNameToId[customFieldname_];

            var set = false;
            
            arr.ForEach(a =>
            {
                var field = a["custom_field_id"].ToString();

                if (field.Equals(customFieldId))
                {
                    ((JValue) a["value"]).Value = value_;
                    set = true;
                }
            });

            if (set) return;
            
            var itemId = itemToken.ExtractString("id");

            var toAdd = new JObject(
                new JProperty("item_id", itemId),
                new JProperty("custom_field_id", customFieldId),
                new JProperty("value", value_)
            );
            
            arr.Add(toAdd);
        }

        private static void setBinLocation(JObject root, string value_)
        {
            if (!SalesBinderAPI.UnitsOfMeasureByName.TryGetValue(value_, out var unitOfMeasure))
            {
                $"Error - can't find a unit of measure with given name {value_}".ConsoleWriteLine();
                return;
            }
            
            setItemDetail(root,InventoryCustomFields.BinLocation,value_);

            if(root["item"].TryExtractToken("unit_of_measure",out var uom))
            {
                uom["id"] = unitOfMeasure.Id;
                uom["full_name"] = unitOfMeasure.LongName;
                uom["short_name"] = unitOfMeasure.ShortName;
            }
        }


        private static (string FieldName, Func<SalesBinderInventoryItem, string> Func, Action<JObject, string> Update)[] _strUpdates =
            new (string,  Func<SalesBinderInventoryItem, string> Func, Action<JObject, string> Update)[]
            {
                (InventoryFields.Name, x => x.Name, (o, s) => ((JValue) o["item"][InventoryFields.Name]).Value = s),
                (InventoryFields.SKU, x => x.SKU, (o, s) => ((JValue) o["item"][InventoryFields.SKU]).Value = s),
                (InventoryFields.BarCode, x => x.BarCode, (o, s) => ((JValue) o["item"][InventoryFields.BarCode]).Value = s),
                (InventoryCustomFields.ProductType, x => x.ProductType, (o, s) => setItemDetail(o, InventoryCustomFields.ProductType, s)),
                (InventoryCustomFields.ProductType2, x => x.ProductType2, (o, s) => setItemDetail(o, InventoryCustomFields.ProductType2, s)),
                (InventoryCustomFields.ProductType3, x => x.ProductType3, (o, s) => setItemDetail(o, InventoryCustomFields.ProductType3, s)),
                (InventoryCustomFields.Clearance, x => x.Clearance, (o, s) => setItemDetail(o, InventoryCustomFields.Clearance, s)),
                (InventoryCustomFields.BinLocation, x => x.BinLocation, (o, s) => setBinLocation(o, s)),
                (InventoryCustomFields.PackSize, x => x.PackSize, (o, s) => setItemDetail(o, InventoryCustomFields.PackSize, s)),
                (InventoryCustomFields.Author, x => x.Author, (o, s) => setItemDetail(o, InventoryCustomFields.Author, s)),
                (InventoryCustomFields.KidsAdult, x => x.KidsOrAdult, (o, s) => setItemDetail(o, InventoryCustomFields.KidsAdult, s)),
                (InventoryCustomFields.FullRrp, x => x.FullRRP, (o, s) => setItemDetail(o, InventoryCustomFields.FullRrp, s)),
                (InventoryCustomFields.Style, x => x.Style, (o, s) => setItemDetail(o, InventoryCustomFields.Style, s)),
                (InventoryCustomFields.Publisher, x => x.Publisher, (o, s) => setItemDetail(o, InventoryCustomFields.Publisher, s)),
                (InventoryCustomFields.MaterialComposition, x => x.MaterialComposition, (o, s) => setItemDetail(o, InventoryCustomFields.MaterialComposition, s)),
                (InventoryCustomFields.CountryOfOrigin, x => x.CountryOfOrigin, (o, s) => setItemDetail(o, InventoryCustomFields.CountryOfOrigin, s)),
                (InventoryCustomFields.CommodityCode, x => x.CommodityCode, (o, s) => setItemDetail(o, InventoryCustomFields.CommodityCode, s)),
                (InventoryCustomFields.Condition, x => x.Condition, (o, s) => setItemDetail(o, InventoryCustomFields.Condition, s)),
                (InventoryCustomFields.VAT, x => x.VAT, (o, s) => setItemDetail(o, InventoryCustomFields.VAT, s)),
                (InventoryCustomFields.SalesRestrictions, x => x.SalesRestrictions, (o, s) => setItemDetail(o, InventoryCustomFields.SalesRestrictions, s)),
            };
        
        private static (string FieldName, Func<SalesBinderInventoryItem, decimal> Func, Action<JObject,decimal> Update)[] _decUpdates = new (string, Func<SalesBinderInventoryItem, decimal> Func, Action<JObject, decimal> Update)[]
        {
            (InventoryFields.Cost, x=>x.Cost,(o,s)=>((JValue) o["item"][InventoryFields.Cost]).Value=s),
            (InventoryFields.Price, x=>x.Price,(o,s)=>((JValue) o["item"][InventoryFields.Price]).Value=s),
        };

        private static (string FieldName, Func<SalesBinderInventoryItem, long> Func, Action<JObject,long> Update)[] _lngUpdates = new (string, Func<SalesBinderInventoryItem, long> Func, Action<JObject, long> Update)[]
        {
            (InventoryFields.Quantity, x=>x.Quantity,(o,s)=>((JValue) o["item"][InventoryFields.Quantity]).Value=s),
        };

        
        public static bool DetectChanges(this SalesBinderInventoryItem potentialUpdates_, IEnumerable<SalesBinderInventoryItem> currentInventory_, bool includeQuantities_, bool sendUpdates_)
        {
            var currentAsDictionary = currentInventory_.ToDictionary(x => x.Id, x => x);
            
            if (!currentAsDictionary.TryGetValue(potentialUpdates_.Id, out var current))
            {
                $"Couldn't find current inventory item to compare update item to. ItemId={potentialUpdates_.Id}"
                    .ConsoleWriteLine();
                return false;
            }

            JObject jsonToUpdate = null;
            var changeFound = false;

            JObject getItemToUpdate() => jsonToUpdate ??
                                         (jsonToUpdate =
                                             SalesBinderAPI.RetrieveJsonForInventoryItem(potentialUpdates_.Id));

            _strUpdates.ForEach(u =>
            {
                var cv = u.Func(current);
                var uv = u.Func(potentialUpdates_);

                if (string.IsNullOrEmpty(cv) && string.IsNullOrEmpty(uv))
                    return;

                if (decimal.TryParse(cv, out var dec1) && decimal.TryParse(uv, out var dec2) && dec1 == dec2)
                    return;
                
                if ( String.CompareOrdinal(cv?.Trim(), uv?.Trim()) != 0)
                {
                    changeFound = true;
                    if(sendUpdates_)
                        u.Update(getItemToUpdate(), uv);
                    else
                        $"\"{current.Name}\": {u.FieldName} field has been changed from '{cv}' to '{uv}'".ConsoleWriteLine();    
                }
            });

            _decUpdates.ForEach(u =>
            {
                var cv = u.Func(current);
                var uv = u.Func(potentialUpdates_);

                if (cv != uv)
                {
                    changeFound = true;
                    if(sendUpdates_)
                        u.Update(getItemToUpdate(), uv);
                    else
                        $"\"{current.Name}\": {u.FieldName} field has been changed from '{cv}' to '{uv}'".ConsoleWriteLine();
                }
            });

            _lngUpdates.ForEach(u =>
            {
                var cv = u.Func(current);
                var uv = u.Func(potentialUpdates_);

                if (cv != uv)
                {
                    if (!includeQuantities_ && u.FieldName.Equals(InventoryFields.Quantity))
                        return;
       
                    changeFound = true;
                    if (sendUpdates_)
                        u.Update(getItemToUpdate(), uv);
                    else
                        $"\"{current.Name}\": {u.FieldName} field has been changed from '{cv}' to '{uv}'".ConsoleWriteLine();
                }
            });

            if (jsonToUpdate != null && sendUpdates_)
            {
                $"Updating '{current.Name}'... ".ConsoleWrite();
                SalesBinderAPI.SendItemUpdate(current.Id,jsonToUpdate.ToString());
            }

            return changeFound;
        }
    }
    
    public class SalesBinderContact
    {
        public static SalesBinderContact Parse(JToken token_)
        {
            return new SalesBinderContact
            {
                FirstName = token_.ExtractString("first_name"),
                LastName = token_.ExtractString("last_name"),
                JobTitle = token_.ExtractString("job_title"),
                Phone = token_.ExtractString("phone"),
                Cell = token_.ExtractString("cell"),
                Email1 = token_.ExtractString("email-1"),
                Email2 = token_.ExtractString("email_2"),
                Modified = token_.ExtractString("modified"),
                Created = token_.ExtractString("created"),
                Id = token_.ExtractString("id")
            };
        }

        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string JobTitle { get; set; }
        public string Phone { get; set; }
        public string Cell { get; set; }
        public string Email1 { get; set; }
        public string Email2 { get; set; }
        public string Id { get; set; }
        public string Modified { get; set; }
        public string Created { get; set; }
    }

    public class SalesBinderAccount
    {
        public static SalesBinderAccount Parse(JToken token_)
        {
            return new SalesBinderAccount
            {
                Name = token_.ExtractString("name"),
                CustomerNumber = token_.ExtractInt("customer_number"),
                OfficeEmail = token_.ExtractString("office_email"),
                OfficePhone = token_.ExtractString("office_phone"),
                Url = token_.ExtractString("url"),
                Id = token_.ExtractString("id"),
                Created = token_.ExtractString("created"),
                Modified = token_.ExtractString("modfified")
            };
        }

        public string Name { get; set; }
        public int CustomerNumber { get; set; }
        public string OfficeEmail { get; set; }
        public string OfficePhone { get; set; }
        public string Url { get; set; }
        public string Id { get; set; }
        public string Created { get; set; }
        public string Modified { get; set; }
    }

    public class SalesBinderInvoice
    {
        public static SalesBinderInvoice Parse(JToken token_)
        {
            var ret = new SalesBinderInvoice
            {
                DocumentNumber = token_.ExtractLong("document_number"),
                Name = token_.ExtractString("name"),
                ContextId = token_.ExtractInt("context_id"),
                TotalCost = token_.ExtractDouble("total_cost"),
                IssueDate = token_.ExtractDateExact("issue_date", "dd/MM/yyyy HH:mm:ss"),
                Created = token_.ExtractDateExact("created", "dd/MM/yyyy HH:mm:ss"),
                Modified = token_.ExtractDateExact("modified", "dd/MM/yyyy HH:mm:ss"),
                CustomerId = token_.ExtractString("customer_id"),
                Id = token_.ExtractString("id"),
            };

            if (ret.IssueDate == DateTime.MinValue)
                ret.IssueDate = token_.ExtractDateExact("issue_date", "dd/MM/yyyy HH:mm:ss");


            var dItems = token_["document_items"] as JArray;

            if (dItems != null)
                ret.Items = dItems.Select(x => SalesBinderInvoiceItem.Parse(x)).ToArray();

            return ret;
        }

        public long DocumentNumber { get; set; }
        public string Name { get; set; }
        public int ContextId { get; set; }
        public double TotalCost { get; set; }
        public DateTime IssueDate { get; set; }
        public string Id { get; set; }
        public string CustomerId { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }

        public SalesBinderInvoiceItem[] Items { get; set; }
    }

    public class SalesBinderInvoiceItem
    {
        public static SalesBinderInvoiceItem Parse(JToken token_)
        {
            return new SalesBinderInvoiceItem
            {
                ItemId = token_.ExtractString("item_id"),
                Cost = token_.ExtractDouble("cost"),
                Quantity = token_.ExtractInt("quantity"),
                Id = token_.ExtractString("id")
            };
        }

        public string ItemId { get; set; }
        public double Cost { get; set; }
        public int Quantity { get; set; }
        public string Id { get; set; }
    }

    public class SalesBinderUnitOfMeasure
    {
        public static SalesBinderUnitOfMeasure Parse(JToken token_)
            => new SalesBinderUnitOfMeasure
            {
                Id = token_.ExtractString("id"),
                ShortName = token_.ExtractString("short_name"),
                LongName = token_.ExtractString("full_name")
            };

        public string Id { get; set; }
        public string ShortName { get; set; }
        public string LongName { get; set; }
    }

    public class StockListOrderer : IComparer<SalesBinderInventoryItem>
    {
        static bool understandClearance(string given_)
        {
            switch (given_.Trim().ToLower())
            {
                case "yes":
                    case "y":
                    case "true":
                    return true;
                default:
                    return false;
            }
        }
        static bool isClearance(SalesBinderInventoryItem item_)
            => string.IsNullOrEmpty(item_?.Clearance) ? false : understandClearance(item_.Clearance);

        static bool? isKids(string given_) =>
            string.IsNullOrEmpty(given_)
                ? default(bool?)
                : "kids".Equals(given_.ToLower());
        
        public int Compare(SalesBinderInventoryItem x, SalesBinderInventoryItem y)
        {
            var xKids = isKids(x.KidsOrAdult);
            var yKids = isKids(y.KidsOrAdult);

            if (xKids.HasValue)
            {
                if (yKids.HasValue)
                {
                    if (yKids.Value != xKids.Value)
                        return xKids.Value.CompareTo(yKids.Value);
                }
                else
                {
                    return 1;
                }
            }
            else if (yKids.HasValue)
            {
                return 1;
            }
            
            
            var xClearance = isClearance(x);
            var yClearnce = isClearance(y);

            if (xClearance != yClearnce)
                return xClearance.CompareTo(yClearnce);

            var xRating = x.Rating ?? 0;
            var yRating = y.Rating ?? 0;

            if (xRating != yRating)
                return xRating.CompareTo(yRating);

            return (x.Name ?? string.Empty).CompareTo(y.Name ?? string.Empty);
        }
    }
}
