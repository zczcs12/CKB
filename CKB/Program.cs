using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using Utility.CommandLine;

namespace CKB
{
    class Program
    {
        [Argument('a',"sbad","Sales Binder Account Download")]
        private static bool SalesBinderAccountsDownload { get; set; }
        
        [Argument('b',"sbal","Sales Binder Account List")]
        private static bool SalesBinderAccountsList { get; set; }
        
        [Argument('c',"sbid","Sales Binder Inventory Download")]
        private static bool SalesBinderInventoryDownload { get; set; }
        
        [Argument('d', "sbil", "Sales Binder Inventory List")]
        private static bool SalesBinderInventoryList { get; set; }
        
        [Argument('e', "sbilc", "Sales Binder Inventory List (Current)")]
        private static bool SalesBinderInventoryListCurrent { get; set; }
        
        [Argument('f',"sbimd", "Sales Binder Image Download")]
        private static bool SalesBinderImageDownload { get; set; }
        
        [Argument('g',"sbcd", "Sales Binder Contact Download")]
        private static bool SalesBinderContactsDownload { get; set; }
        
        [Argument('h',"sbcl","Sales Binder Contact List")]
        private static bool SalesBinderContactsList { get; set; }
        
        [Argument('i',"sbinvd","Sales Binder Invoices Download")]
        private static bool SalesBinderInvoicesDownload { get; set; }
        
        [Argument('v',"sbsd","Sales Binder Settings Download")]
        private static bool SalesBinderSettingsDownload { get; set; }

        [Argument('x',"sbld","Sales Binder Locations Download")]
        private static bool SalesBinderLocationsDownload { get; set; }
        
        [Argument('y',"sbuqf","Sales Binder Update Quantities from csv file")]
        private static bool SalesBinderUpdateQuantitiesFromFile { get; set; }

        [Argument('j',"kle","Keepa lookup ensure. Lookup given IDs in local keepa records. Try to update any missing.")]
        private static bool KeepaLookupEnsure { get; set; }
        
        [Argument('k',"klf","Keepa lookup force i.e. force refresh of record for these identifiers.")]
        private static bool KeepaLookupForceRefresh { get; set; }
        
        [Argument('l',"klri","Keepa lookup refresh inventory - if not updated in the last 24 hours")]
        private static bool KeepaLookupRefreshCurrentInventory { get; set; }
        
        [Argument('m',"gwl","Generate warehouse list xlsx from given csv (id,qty) file")]
        private static bool GenerateListForWarehouse { get; set; }
        
        [Argument('n', "gsr", "Generate sales report csv")]
        private static bool GenerateSalesReport { get; set; }
        
        [Argument('o', "gir", "Generate current inventory csv")]
        private static bool GenerateCurrentInventoryReport { get; set; }
        
        [Argument('p',"gslf","Generate stock list from file")]
        private static bool GenerateStockListFromFile { get; set; }
        
        [Argument('q',"kru","Try to topup keepa records")]
        private static bool TopupKeepaRecords { get; set; }
        
        [Argument('t', "test", "Test code")]
        private static bool Test { get; set; }
        
        [Argument('z',"output","File to write results to")]
        private static bool OutputFilePath { get; set; }

        static void Main(string[] args)
        {
            Arguments.Populate();

            if (args == null || args.Length == 0 || args.First().Equals("/?"))
            {
                Arguments.GetArgumentInfo()
                    .Select(x => new[] {$"--{x.LongName}", $"{x.HelpText}"})
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
                SalesBinderAPI.RetrieveAndSaveInventory();
            if(SalesBinderImageDownload)
                SalesBinderAPI.DownloadBookImages();
            if(SalesBinderContactsDownload)
                SalesBinderAPI.RetrieveAndSaveContacts();
            if(SalesBinderAccountsDownload)
                SalesBinderAPI.RetrieveAndSaveAccounts();
            if(SalesBinderInvoicesDownload)
                SalesBinderAPI.RetrieveAndSaveInvoices();
            if(SalesBinderSettingsDownload)
                SalesBinderAPI.RetrieveAndSaveSettings();
            if(SalesBinderLocationsDownload)
                SalesBinderAPI.RetrieveAndSaveLocations();
            
            if (SalesBinderInventoryList || SalesBinderInventoryListCurrent)
            {
                var list = SalesBinderAPI.Inventory;

                if (SalesBinderInventoryListCurrent)
                    list = list.Where(x => x.Quantity > 0).ToArray();
                    
                (list.Count()==0 ? "No local inventory found" : list.FormatIntoColumnsReflectOnType(60))
                    .ConsoleWriteLine();
            }

            if (SalesBinderAccountsList)
            {
                var list = SalesBinderAPI.Accounts;
                
                (list.Length==0 ? "No local accounts found" : list.FormatIntoColumnsReflectOnType(100))
                    .ConsoleWriteLine();
            }

            if (KeepaLookupEnsure || KeepaLookupForceRefresh)
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
                    
                    var result = KeepaAPI.GetDetailsForIdentifiers(identifiers,forceRefresh_:KeepaLookupForceRefresh);

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

            if (GenerateListForWarehouse)
            {
                if (OutputFilePath == false || !hasArgument("output"))
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
                        .WriteWarehouseFile(getArgument("output"));
                }
            }

            if (GenerateSalesReport)
            {
                if (OutputFilePath == false || !hasArgument("output"))
                    "You need to supply an output filepath (which should be an csv) to write the list to".ConsoleWriteLine();
                else if (SalesBinderAPI.Invoices == null)
                    "No invoices found.  Have you downloaded them yet?".ConsoleWriteLine();
                else if (SalesBinderAPI.Inventory == null)
                    "No inventory found.  Have you downloaded them yet?".ConsoleWriteLine();
                else if(SalesBinderAPI.Accounts==null)
                    "No accounts found.  Have you downloaded them yet?".ConsoleWriteLine();
                else if (OutputFilePath == false || !hasArgument("output"))
                    "Please specify an output file path, which should be a csv".ConsoleWriteLine();
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
                    File.WriteAllText(getArgument("output"),sb.ToString());
                }
            }

            if (GenerateCurrentInventoryReport)
            {
                if (OutputFilePath == false || !hasArgument("output"))
                    "You need to supply an output filepath (which should be an csv) to write the list to".ConsoleWriteLine();
                else if (SalesBinderAPI.Inventory == null)
                    "No inventory found.  Have you downloaded them yet?".ConsoleWriteLine();
                else
                {
                    var setups = new (string Title, Func<SalesBinderInventoryItem, string> Func)[]
                    {
                        ("Name", x => x.Name),
                        ("Author", x => x.Author),
                        ("Qty", x => $"{x.Quantity}"),
                        ("Cost", x => $"{x.Cost}"),
                        ("Price", x => $"{x.Price}"),
                        ("SKU", x => x.SKU),
                        ("ItemNumber", x => $"{x.ItemNumber}"),
                        ("BarCode", x => x.BarCode),
                        ("Style", x => x.Style),
                        ("KidsOrAdult", x => x.KidsOrAdult),
                        ("ProductType", x => x.ProductType),
                        ("BinLocation", x => x.BinLocation),
                        ("PackSize", x => x.PackSize),
                        ("VAT", x => x.VAT),
                    };
                    
                    var sb = new StringBuilder();
                    sb.AppendLine(setups.Select(s => s.Title).ToArray().Join(","));
                    SalesBinderAPI.Inventory
                        .Where(x=>x.Quantity>0)
                        .ForEach(row=>sb.AppendLine(setups.Select(s=>encaseStringWithComma(s.Func(row))).ToArray().Join(",")));
                    File.WriteAllText(getArgument("output"),sb.ToString());
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

            if (GenerateStockListFromFile)
            {
                if (OutputFilePath == false || !hasArgument("output"))
                {
                    "You need to supply an output filepath (which should be an xlsx) to write the list to".ConsoleWriteLine();
                }
                else if (!lookup.ArgumentDictionary.TryGetValue("gslf", out var arg) || arg==null || string.IsNullOrEmpty(arg.ToString()))
                    "You need to supply a csv argument of the identifiers you want to lookup".ConsoleWriteLine();
                else if (!File.Exists(arg.ToString()))
                    $"File '{arg.ToString()}' does not exist.".ConsoleWriteLine();
                else
                {
                    var items = File.ReadAllLines(arg.ToString()).Distinct().ToList();

                    var imageLookup = items.Select(x =>
                        {
                            SalesBinderAPI.InventoryByBarcode.TryGetValue(x, out var book);
                            return (Identifier: x, ImagePath: ExtensionMethods.FindImagePath(x), Book: book);
                        })
                        .ToList();
                    
                    imageLookup.Where(x=>string.IsNullOrEmpty(x.ImagePath))
                        .ForEach(x=>$"Couldn't find image for {x.Identifier}".ConsoleWriteLine());
                    
                    imageLookup.WriteStockListFile(getArgument("output"));
                }
            }

            if (SalesBinderUpdateQuantitiesFromFile)
            {
                if (!lookup.ArgumentDictionary.TryGetValue("sbuqf", out var arg) || arg==null || string.IsNullOrEmpty(arg.ToString()))
                    "You need to supply a csv argument of the identifiers you want to lookup".ConsoleWriteLine();
                else if (!File.Exists(arg.ToString()))
                    $"File '{arg.ToString()}' does not exist.".ConsoleWriteLine();
                else
                {
                    var inventoryByBarCode = SalesBinderAPI.InventoryByBarcode;

                    var contents = File.ReadAllLines(args.ToString())
                        .Select(x => x.Split(','))
                        .Select(x =>
                        {
                            if (x.Length != 2)
                                return (Success: false, BarCode: null as string, NewQuantity: default(int),
                                    Error: "Each line needs to have two entries, first barcode, then new quantity");

                            var barcode = x[0];
                            if (!inventoryByBarCode.TryGetValue(barcode, out var inventoryItem))
                                return (Success: false, BarCode: barcode, NewQuantity: default(int),
                                    Error: $"Could find an item in local inventory records with barcode {barcode}");

                            if (!int.TryParse(x[1], out var quantity))
                                return (Success: false, BarCode: barcode, NewQuantity: default(int),
                                    Error: $"Couldn't parse given quantity ('{x[1]}') to an integer");

                            return (Success: true, BarCode: barcode, NewQuantity: quantity,
                                Error: null);
                        })
                        .ToArray();

                    if (contents.Any(x => x.Success == false))
                    {
                        $"There were some problems with csv fields provided ('{arg.ToString()}'):"
                            .ConsoleWriteLine();
                        
                        contents.Where(x=>x.Success==false)
                            .Select(x=>new[] {x.BarCode,x.Error})
                            .FormatIntoColumns(new[] {"Barcode","Error"})
                            .ConsoleWriteLine();
                    }
                    else
                    {
                        contents.ForEach(x=>SalesBinderAPI.UpdateQuantity(x.BarCode,x.NewQuantity));   
                    }
                }
            }

            if (Test)
            {
                // SalesBinderAPI.Inventory.GroupBy(x=>x.BarCode)
                //     .Where(x=>x.Count()>1)
                //     .SelectMany(x=>x)
                //     .Select(x=>new [] {$"{x.Name}",$"{x.BarCode}",$"{x.Id}"})
                //     .FormatIntoColumns(new[] {"Name","BarCode","ItemId"})
                //     .ConsoleWriteLine();
                // SalesBinderAPI.UpdateQuantity("9781786484734", 1);
                // SalesBinderAPI.UnitsOfMeasure.ForEach(x=>$"{x.Key} = {x.Value}".ConsoleWriteLine());

                SalesBinderAPI.Inventory.Where(x => x.BarCode.Length == 13 && x.Quantity > 0)
                    .ForEach(x=>BlackwellsAPI.TryGetImageForIdentifier(x.BarCode));
            }
        }
    }
}
