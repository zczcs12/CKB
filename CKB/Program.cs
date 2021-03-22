using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using Utility.CommandLine;

namespace CKB
{
    class Program
    {
        [Argument('a',"sbad","Sales Binder Account Download")]
        private static bool SalesBinderAccountsDownload { get; set; }
        
        // [Argument('b',"sbal","Sales Binder Account List => csv")]
        private static bool SalesBinderAccountsList { get; set; }
        
        [Argument('c',"sbid","Sales Binder Inventory Download")]
        private static bool SalesBinderInventoryDownload { get; set; }
        
        [Argument('d', "sbil", "Sales Binder Inventory List => csv")]
        private static bool SalesBinderInventoryList { get; set; }
        
        [Argument('e', "sbilc", "Sales Binder Inventory List (Current) => csv")]
        private static bool SalesBinderInventoryListCurrent { get; set; }
        
        [Argument('f',"sbimd", "Sales Binder Image Download")]
        private static bool SalesBinderImageDownload { get; set; }
        
        [Argument('g',"sbcd", "Sales Binder Contact Download")]
        private static bool SalesBinderContactsDownload { get; set; }
        
        [Argument('i',"sbinvd","Sales Binder Invoices Download")]
        private static bool SalesBinderInvoicesDownload { get; set; }
        
        [Argument('j',"sbiu","Sales Binder Inventory Update (from file)")]
        private static bool SalesBinderInventoryUpdate { get; set; }
        
        // [Argument('v',"sbsd","Sales Binder Settings Download")]
        private static bool SalesBinderSettingsDownload { get; set; }

        // [Argument('x',"sbld","Sales Binder Locations Download")]
        private static bool SalesBinderLocationsDownload { get; set; }
        
        [Argument('u', "sbimu","Sales Binder Image Upload (find for inventory without one)")]
        private static bool SalesBinderFindAndUploadImagesForInventoryWithoutAnImage { get; set; }

        [Argument('j',"kle","Keepa lookup ensure. Lookup given IDs in local keepa records. Try to update any missing.")]
        private static bool KeepaLookupEnsure { get; set; }
        
        [Argument('k',"force","Force updates")]
        private static bool Force { get; set; }
        
        [Argument('/',"klp","Keepa lookup prime records.  Try to get a keepa record for anthing we've not tried before.")]
        private static bool KeepaLookupPrimeRecords { get; set; }
        
        [Argument('l',"klri","Keepa lookup refresh inventory - if not updated in the last 24 hours")]
        private static bool KeepaLookupRefreshCurrentInventory { get; set; }
        
        // [Argument('m',"gwl","Generate warehouse list xlsx from given csv (id,qty) file")]
        private static bool GenerateListForWarehouse { get; set; }
        
        [Argument('n', "gsr", "Generate sales report csv")]
        private static bool GenerateSalesReport { get; set; }
        
        [Argument('q',"kru","Try to topup keepa records")]
        private static bool TopupKeepaRecords { get; set; }
        
        [Argument('t', "test", "Test code")]
        private static bool Test { get; set; }
        
        [Argument('#',"gsli","Generate stock list from inventory => xlsx")]
        private static bool GenerateStockListFromInventory { get; set; }

        static void Main(string[] args)
        {
            Arguments.Populate();

            if (EnvironmentSetup.IsProperlySetup(out var error) == false)
            {
                error.ConsoleWriteLine();
                return;
            }

            if (args == null || args.Length == 0 || args.First().Equals("/?") || args.Any(a=>a.StartsWith("-") && !a.StartsWith("--")))
            {
                Arguments.GetArgumentInfo()
                    .Select(x => new[] {$"--{x.LongName}", $"{x.HelpText}"})
                    .OrderBy(x=>x[0])
                    .FormatIntoColumns(new[] {"Argument", "Help"})
                    .ConsoleWriteLine();
                return;
            }

            var lookup = Arguments.Parse(string.Join(" ", args));

            bool hasArgument(string key_) => lookup.ArgumentDictionary.ContainsKey(key_) &&
                                             !string.IsNullOrEmpty($"{lookup.ArgumentDictionary[key_]}");

            string getArgument(string key_) => hasArgument(key_) ? $"{lookup.ArgumentDictionary[key_]}" : null;
            
            string encaseStringWithComma(string input_) => !string.IsNullOrEmpty(input_) && input_.Contains(',')
                ? $"\"{input_}\""
                : input_;
            
            if(SalesBinderInventoryDownload)
                SalesBinderAPI.RetrieveAndSaveInventory(topup_:!Force);
            if(SalesBinderImageDownload)
                SalesBinderAPI.DownloadBookImages();
            if(SalesBinderContactsDownload)
                SalesBinderAPI.RetrieveAndSaveContacts();
            if(SalesBinderAccountsDownload)
                SalesBinderAPI.RetrieveAndSaveAccounts();
            if(SalesBinderInvoicesDownload)
                SalesBinderAPI.RetrieveAndSaveInvoices(topup_:!Force);
            if(SalesBinderSettingsDownload)
                SalesBinderAPI.RetrieveAndSaveSettings();
            if(SalesBinderLocationsDownload)
                SalesBinderAPI.RetrieveAndSaveLocations();

            new (bool Do, string Arg, bool OnlyCurrent)[]
                {
                    (SalesBinderInventoryList, "sbil", false),
                    (SalesBinderInventoryListCurrent, "sbilc", true)
                }
                .Where(x => x.Do)
                .ForEach(set =>
                {
                    var outputFilePath = getArgument(set.Arg);

                    if (string.IsNullOrEmpty(outputFilePath))
                    {
                        $"You need to supply an output filepath after '--{set.Arg}' (which should be an csv) to write the list to"
                            .ConsoleWriteLine();
                        return;
                    }

                    var list = SalesBinderAPI.RetrieveAndSaveInventory(true);

                    if (set.OnlyCurrent)
                        list = list.Where(x => x.Quantity > 0).ToArray();

                    var props = typeof(SalesBinderInventoryItem).GetProperties(BindingFlags.Instance |
                                                                               BindingFlags.GetProperty |
                                                                               BindingFlags.Public)
                        .Where(x => !x.Name.Contains("Image"))
                        .ToArray();

                    var sb = new StringBuilder();
                    sb.AppendLine(string.Join(",", props.Select(p => p.Name)));

                    using (var writer = new StreamWriter(outputFilePath))
                    using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                    {
                        csv.WriteRecords(list
                            .OrderByDescending(x => Math.Abs(x.Quantity))
                            .ThenBy(x => x.Name));
                    }
                });
                
            
            if (SalesBinderAccountsList)
            {
                var list = SalesBinderAPI.Accounts;
                
                (list.Length==0 ? "No local accounts found" : list.FormatIntoColumnsReflectOnType(100))
                    .ConsoleWriteLine();
            }

            if (KeepaLookupEnsure)
            {
                var key = KeepaLookupEnsure ? "kle" : "klf";
                if (!lookup.ArgumentDictionary.TryGetValue(key, out var arg) || arg==null || string.IsNullOrEmpty(arg.ToString()))
                    "You need to supply a csv argument of the identifiers you want to lookup".ConsoleWriteLine();
                else
                {
                    string[] identifiers;
                    if (File.Exists(arg.ToString()))
                        identifiers = File.ReadAllLines(arg.ToString()).Distinct().ToArray();
                    else
                        identifiers = arg.ToString().Split(',');
                    
                    var result = KeepaAPI.GetDetailsForIdentifiers(identifiers,forceRefresh_:Force);

                    var failed = result.Where(x => x.Value == null);
                    if (failed.Any())
                    {
                        $"{failed.Count()} items failed:".ConsoleWriteLine();
                        failed.Select(x => new[] {x.Key})
                            .FormatIntoColumns(new[] {"Failed identifier"})
                            .ConsoleWriteLine();
                    }

                    var succeeded = result.Where(x => x.Value != null);
                    if(succeeded.Any())
                        succeeded.Select(x=>x.Value)
                            .Select(x=>new [] {x.Asin,x.Title,x.Author,x.Manufacturer,x.Binding,x.ImagesCSV,x.EanList==null ? string.Empty : string.Join(",",x.EanList), x.CategoryTree==null ? string.Empty : string.Join(" / ",x.CategoryTree)})
                            .FormatIntoColumns(new[] {"Asin","Title","Author","Manufacturer","Binding","ImagesCSV","EanList","CategoryTree"},100)
                            .ConsoleWriteLine();
                }
            }

            if (KeepaLookupRefreshCurrentInventory)
            {
                var items = SalesBinderAPI.Inventory.Where(x => x.Quantity > 0 && !string.IsNullOrEmpty(x.BarCode))
                    .Where(x => KeepaAPI.LastLookupTime(x.BarCode).HasValue == false
                                || (DateTime.Now - KeepaAPI.LastLookupTime(x.BarCode).Value).TotalDays < 1.0);
                
                // only want to do 100 as don't want to blow limits on keepa
                if(!items.Any())
                    "Nothing to do - have been updated or attempted to be updated at least oncein the last 24 hours"
                        .ConsoleWriteLine();
                else
                {
                    var ids = items.Take(100).Select(x => x.BarCode).Distinct().ToArray();
                    KeepaAPI.GetDetailsForIdentifiers(ids, forceRefresh_: true);
                }
            }

            if (KeepaLookupPrimeRecords)
            {
                var items = SalesBinderAPI.Inventory.Where(x => !string.IsNullOrEmpty(x.BarCode))
                    .Where(x => KeepaAPI.LastLookupTime(x.BarCode) == null);
                
                if(!items.Any())
                    $"Have already tried to get keepa records for all items in inventory".ConsoleWriteLine();
                else
                {
                    var ids = items.Take(100).Select(x => x.BarCode).Distinct().ToArray();
                    KeepaAPI.GetDetailsForIdentifiers(ids);
                }
            }

            if (GenerateListForWarehouse)
            {
                if (!hasArgument("gwl"))
                {
                    "You need to supply an output filepath (which should be an xlsx) to write the list to".ConsoleWriteLine();
                }
                else if (!lookup.ArgumentDictionary.TryGetValue("gwl", out var filename) || string.IsNullOrEmpty(filename?.ToString()))
                {
                    "No file name provided for generating warehouse list".ConsoleWriteLine();
                }
                else if (!File.Exists(filename.ToString()))
                {
                    $"Could not find file '{filename.ToString()}' as source for generating warehouse list".ConsoleWriteLine();
                }
                else
                {
                    filename.ToString().ConsoleWriteLine();
                    var lines = File.ReadLines(filename.ToString())
                        .Select(x => x.Split(','))
                        .Select(x => (Id: x[0].Replace("-",string.Empty), Quantity: int.Parse(x[1])))
                        .ToList();
                    
                    var records = KeepaAPI.GetDetailsForIdentifiers(lines.Select(x => x.Id).ToArray());
                    
                    lines.Select(x =>
                    {
                        records.TryGetValue(x.Id, out var rec);
                        return (x.Id, x.Quantity, rec);
                    })
                        .WriteWarehouseFile(getArgument("gwl"));
                }
            }

            if (GenerateSalesReport)
            {
                if (!hasArgument("gsr"))
                    "You need to supply an output filepath (which should be an csv) to write the list to".ConsoleWriteLine();
                else if (SalesBinderAPI.Invoices == null)
                    "No invoices found.  Have you downloaded them yet?".ConsoleWriteLine();
                else if (SalesBinderAPI.Inventory == null)
                    "No inventory found.  Have you downloaded them yet?".ConsoleWriteLine();
                else if(SalesBinderAPI.Accounts==null)
                    "No accounts found.  Have you downloaded them yet?".ConsoleWriteLine();
                else
                {
                    var invoicesByAllItems = SalesBinderAPI.Invoices
                        .SelectMany(inv => inv.Items.Select(i => (Invoice: inv, Item: i)));

                    var inventoryLookup = SalesBinderAPI.InventoryById;
                    var accountLookup = SalesBinderAPI.AccountById;

                    var rows = invoicesByAllItems.Select(i =>
                        {
                            inventoryLookup.TryGetValue(i.Item.ItemId, out var book);
                            accountLookup.TryGetValue(i.Invoice.CustomerId, out var account);
                            var keepaRecord = book == null ? null : KeepaAPI.GetRecordForIdentifier(book.BarCode);

                            return (Invoice: i.Invoice, SalesItem: i.Item, Book: book, Account: account, Keepa:keepaRecord);
                        })
                        .ToList();

                    string getItemAt(string[] arr_, int index_) =>
                        arr_ == null || arr_.Length < (index_ + 1) ? null : arr_[index_];

                    var setups =
                        new (string Title, Func<(SalesBinderInvoice Invoice, SalesBinderInvoiceItem SalesItem,
                            SalesBinderInventoryItem Book, SalesBinderAccount Account, KeepaRecord Keepa), string> Func)[]
                        {
                            ("i.Date", x => $"{x.Invoice.IssueDate:yyyyMMdd}"),
                            ("i.Number", x => $"{x.Invoice.DocumentNumber}"),
                            ("i.TotalCost", x => $"{x.Invoice.TotalCost}"),
                            ("i.ItemCost", x => $"{x.SalesItem.Cost}"),
                            ("i.ItemQty", x => $"{x.SalesItem.Quantity}"),
                            ("account", x => x.Account?.Name),
                            ("sb.kidsOrAdult",x => x.Book?.KidsOrAdult),
                            ("sb.price",x => $"{x.Book?.Price}"),
                            ("sb.productType",x => $"{x.Book?.ProductType}"),
                            ("sb.author",x => $"{x.Book?.Author}"),
                            ("sb.packSize",x => $"{x.Book?.PackSize}"),
                            ("sb.vat",x => $"{x.Book?.VAT}"),
                            ("k.amznCat1", x => getItemAt(x.Keepa?.CategoryTree, 0)),
                            ("k.amznCat2", x => getItemAt(x.Keepa?.CategoryTree, 1)),
                            ("k.amznCat3", x => getItemAt(x.Keepa?.CategoryTree, 2)),
                            ("k.amznCat4", x => getItemAt(x.Keepa?.CategoryTree, 3)),
                            ("k.amznCat5", x => getItemAt(x.Keepa?.CategoryTree, 4)),
                            ("k.amznCat6", x => getItemAt(x.Keepa?.CategoryTree, 5)),
                            ("k.amznCat7", x => getItemAt(x.Keepa?.CategoryTree, 6)),
                        };
                    
                    var sb = new StringBuilder();
                    sb.AppendLine(setups.Select(s => s.Title).ToArray().Join(","));
                    rows.ForEach(row=>sb.AppendLine(setups.Select(s=>encaseStringWithComma(s.Func(row))).ToArray().Join(",")));
                    File.WriteAllText(getArgument("gsr"),sb.ToString());
                }
            }

            if (TopupKeepaRecords)
            {
                var missing = SalesBinderAPI.Inventory
                    .Select(x => (Inventory: x, HaveKeepa: KeepaAPI.HaveLocalRecord(x.BarCode)))
                    .Where(x => x.HaveKeepa==false)
                    .Where(x=>!KeepaAPI.LastLookupTime(x.Inventory.BarCode).HasValue==false)
                    .ToList();

                if (missing.Any())
                {
                    KeepaAPI.GetDetailsForIdentifiers(missing.Take(100).Select(x => x.Inventory.BarCode).Distinct().ToArray());
                }
            }

            if (SalesBinderFindAndUploadImagesForInventoryWithoutAnImage)
            {
                SalesBinderAPI.Inventory.Where(x=>x.HasImageSaved()==false && x.Quantity>0)
                    .Select(x=>(Item:x,Image:ExtensionMethods.FindImagePath(x.BarCode)))
                    .ForEach(x=>
                    {
                        $"{x.Item.BarCode} : found image : {x.Image}".ConsoleWriteLine();
                
                        if(!string.IsNullOrEmpty(x.Image))
                            SalesBinderAPI.UploadImage(x.Item.BarCode,x.Image);
                    });
                
            }
            
            if (GenerateStockListFromInventory)
            {
                if (!hasArgument("gsli"))
                {
                    "You need to supply an output filepath (which should be an xlsx) to write the list to".ConsoleWriteLine();
                }
                else
                {
                    var inventory = SalesBinderAPI.RetrieveAndSaveInventory(true);

                    if (!inventory.Any())
                    {
                        $"There is no salesbinder inventory locally.  run 'CKB.exe --sbid' to download inventory before trying to generate a stock list"
                            .ConsoleWriteLine();
                    }
                    else
                    {
                        inventory.Where(x=>x.Quantity>0)
                            .Select(x => (BarCode:x.BarCode, Image:ExtensionMethods.FindImagePath(x.BarCode), Item:x))
                            .OrderByDescending(x=>x.Item,new StockListOrderer())
                            .WriteStockListFile(getArgument("gsli"));
                    }
                }
            }


            if (SalesBinderInventoryUpdate)
            {
                if (!hasArgument("gsiu"))
                {
                    "You need to supply an argument to --gsiu which is the path to the csv file that containst the updates".ConsoleWriteLine();
                }
                else
                {
                    SalesBinderInventoryItem[] recs;

                    using (var reader = new StreamReader(getArgument("gsiu")))
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        recs = csv.GetRecords<SalesBinderInventoryItem>().ToArray();
                    }

                    recs.ForEach(potentialUpdate => { potentialUpdate.DetectChanges(Force); });
                }
            }
            
            if (Test)
            {
                 var c = SalesBinderAPI.RetrieveJsonForInventoryItem("5b460759-5764-4a88-b3c5-2d8c3f71d8bf");
                 c.ToString().ConsoleWriteLine();
                 File.WriteAllText(@"e:\temp.json",c.ToString());
            }
        }
    }
}
