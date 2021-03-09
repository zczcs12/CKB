using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Web.Hosting;
using DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Path = System.IO.Path;

namespace CKB
{
    internal static class SalesBinderAPI
    {

        private const string URI_ROOT = "https://ckb.salesbinder.com/api/2.0/";
        public static readonly  string SAVE_PATH_ROOT = $@"{Environment.GetEnvironmentVariable("CKB_DATA_ROOT",EnvironmentVariableTarget.User)}\SalesBinder";

        private static string InventoryFilePath => $@"{SAVE_PATH_ROOT}\inventory.json";
        private static string ContactsFilePath = $@"{SAVE_PATH_ROOT}\contacts.json";
        private static string AccountsFilePath = $@"{SAVE_PATH_ROOT}\accounts.json";
        private static string InvoicesFilePath = $@"{SAVE_PATH_ROOT}\invoices.json";
        private static string SettingsFilePath = $@"{SAVE_PATH_ROOT}\settings.json";
        private static string LocationsFilePath = $@"{SAVE_PATH_ROOT}\locations.json";
        private static string UnitsOfMeasureFilePath = $@"{SAVE_PATH_ROOT}\units_of_measure.json";

        #region Inventory
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
                    .ToDictionary(x => x.Key, x => x.OrderByDescending(b=>b.Quantity).First())
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
        
        private static readonly Lazy<SalesBinderInvoice[]> _invoices = new Lazy<SalesBinderInvoice[]>(()=>
            File.Exists(InvoicesFilePath)
                ? JsonConvert.DeserializeObject<SalesBinderInvoice[]>(File.ReadAllText(InvoicesFilePath))
                : null);

        public static SalesBinderInvoice[] Invoices => _invoices.Value;
        
        #endregion
        
        #region Units of Measure

        private static readonly Lazy<SalesBinderUnitOfMeasure[]> _unitsOfMeasure = new Lazy<SalesBinderUnitOfMeasure[]>(
            () =>
            {
                var s = JToken.Parse(File.ReadAllText(UnitsOfMeasureFilePath));
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

        private static readonly Lazy<HttpClient> _client = new Lazy<HttpClient>(() =>
        {
            var apiKey = Environment.GetEnvironmentVariable("SALESBINDER_API_KEY", EnvironmentVariableTarget.User);

            var client = new HttpClient();
            var byteArray = Encoding.ASCII.GetBytes($"{apiKey}:x");
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            return client;
        });

        private static Lazy<WebClient> _webClient = new Lazy<WebClient>(() => new WebClient());


        public static void RetrieveAndSaveInventory()
            => retrieveAndSave(RetrieveInventory, InventoryFilePath);

        public static IEnumerable<SalesBinderInventoryItem> RetrieveInventory()
            => retrievePaginated("items", "items", SalesBinderInventoryItem.Parse,"&pageLimit=100");
        
        private static IEnumerable<T> retrievePaginated<T>(string apiName_, string groupTokenName_,
            Func<JToken, T> creator_, string optionalArgs_=null)
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

        public static IEnumerable<SalesBinderInvoice> RetrieveInvoices()
            => retrievePaginated("documents", "documents", x => SalesBinderInvoice.Parse(x), "&contextId=5&pageLimit=200");

        public static void RetrieveAndSaveInvoices()
            => retrieveAndSave(RetrieveInvoices, InvoicesFilePath);
        
        private static void retrieveAndSave<T>(Func<IEnumerable<T>> func_, string filePath_)
        {
            var allItems = func_();
            var txt = JsonConvert.SerializeObject(allItems.ToArray(), Formatting.None);
            File.WriteAllText(filePath_, txt);
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
                        (item,item.ImageURLSmall,item.ImageFilePath(ImageSize.Small)),
                        (item,item.ImageURLMedium,item.ImageFilePath(ImageSize.Medium)),
                        (item,item.ImageURLLarge,item.ImageFilePath(ImageSize.Large)),
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

        private static string putJson(string urlEnding_, string json_)
        {
            blockIfNecessary();
            var httpContent = new StringContent(json_, Encoding.UTF8, "application/json");
            var uri = new Uri($"{URI_ROOT}{urlEnding_}");
            var response = _client.Value.PutAsync(uri, httpContent).Result;
            _lastRequestTime=DateTime.Now;
            return response?.Content.ReadAsStringAsync().Result;
        }

        private static string postImage(string urlEnding_, string filePath_)
        {
            blockIfNecessary();
            var uri = new Uri($"{URI_ROOT}{urlEnding_}");

            var request = new HttpRequestMessage(HttpMethod.Post, uri);
            request.Content = new ByteArrayContent(File.ReadAllBytes(filePath_));
            request.Content.Headers.ContentType=MediaTypeHeaderValue.Parse("image/jpeg");

            var response = _client.Value.SendAsync(request).Result;
            _lastRequestTime=DateTime.Now;
            
            return response?.Content.ReadAsStringAsync().Result;
        }
        
        
        public static void UpdateQuantity(string identifier_, long newQuantity_)
        {
            if (!_inventoryByBarcode.Value.TryGetValue(identifier_, out var record))
            {
                $"Don't know of an inventory item with barcode {identifier_} so can't update it."
                    .ConsoleWriteLine();
                return;
            }

            var urlEnding = $"items/{record.Id}.json";
            
            var json = "{ \"item\": { \"quantity\" : <qty> } }"
                .Replace("<qty>", $"{newQuantity_}");

            var response = putJson(urlEnding, json);

            response.ConsoleWriteLine();
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
        Small, Medium, Large
    }

    public class SalesBinderInventoryItem
    {
        public static SalesBinderInventoryItem Parse(JToken token_)
        {
            var ret = new SalesBinderInventoryItem
            {
                Name = token_.ExtractString("name"),
                Quantity = token_.ExtractLong("quantity"),
                Cost = token_.ExtractDecimal("cost"),
                Price = token_.ExtractDecimal("price"),
                SKU = token_.ExtractString("sku"),
                ItemNumber = token_.ExtractLong("item_number"),
                BarCode = token_.ExtractString("barcode"),
                Id = token_.ExtractString("id"),
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
            
            var details = token_["item_details"] as JArray;

            foreach (var det in details)
            {
                var field = det["custom_field"]["name"].ToString();
                var val = det["value"].ToString();

                if ("Style".Equals(field))
                    ret.Style = val;
                else if ("Country of Origin".Equals(field))
                    ret.CountryOfOrigin = val;
                else if ("Publisher".Equals(field))
                    ret.Publisher = val;
                else if ("Kids/adult".Equals(field))
                    ret.KidsOrAdult = val;
                else if ("Product Type".Equals(field))
                    ret.ProductType = val;
                else if ("Author".Equals(field))
                    ret.Author = val;
                // else if ("Bin location".Equals(field))
                //     ret.BinLocation = val;
                else if ("Pack size".Equals(field))
                    ret.PackSize = val;
                else if ("VAT".Equals(field))
                    ret.VAT = val;
            }

            var images = token_["images"];

            ret.ImageURLSmall = images?.First.ExtractString("url_small");
            ret.ImageURLMedium = images?.First.ExtractString("url_medium");
            ret.ImageURLLarge = images?.First.ExtractString("url_large");

            return ret;
        }

        public string Name { get; set; }
        public long Quantity { get; set; }
        public decimal Cost { get; set; }
        public decimal Price { get; set; }
        public string SKU { get; set; }
        public long ItemNumber { get; set; }
        public string BarCode { get; set; }

        public string Id { get; set; }
        
        public string Style { get; set; }
        public string CountryOfOrigin { get; set; }
        public string Publisher { get; set; }
        public string KidsOrAdult { get; set; }
        public string ProductType { get; set; }
        public string BinLocation { get; set; }
        
        public string BinLocationId { get; set; }
        public string PackSize { get; set; }
        public string Author { get; set; }
        public string VAT { get; set; }

        public string ImageURLSmall { get; set; }
        public string ImageURLMedium { get; set; }
        public string ImageURLLarge { get; set; }
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
                Id=token_.ExtractString("id")
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
                IssueDate = token_.ExtractDate("issue_date"),
                Created = token_.ExtractDate("created"),
                Modified = token_.ExtractDate("modified"),
                CustomerId = token_.ExtractString("customer_id"),
                Id=token_.ExtractString("id"),
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
    
}
