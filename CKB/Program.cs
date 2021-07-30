using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using CsvHelper;
using Utility.CommandLine;

namespace CKB
{
    class Program
    {
        static void Main(string[] args)
        {
            if (EnvironmentSetup.IsProperlySetup(out var error) == false)
            {
                error.ConsoleWriteLine();
                return;
            }

            if (args == null || args.Length == 0 || args.First().Equals("/?") || args.Any(a=>a.StartsWith("-") && !a.StartsWith("--")))
            {
                Arguments.GetArgumentInfo(typeof(Args))
                    .Select(x => new[] {$"--{x.LongName}", $"{x.HelpText}"})
                    .OrderBy(x=>x[0])
                    .FormatIntoColumns(new[] {"Argument", "Help"})
                    .ConsoleWriteLine();
                return;
            }

            Arguments.Populate(typeof(Args));
            var arguments = Arguments.Parse(string.Join(" ", args),typeof(Args));

            if(Args.SalesBinderInventoryDownload)
                SalesBinderAPI.RetrieveAndSaveInventory(topup_:!Args.Force);
            if(Args.SalesBinderImageDownload)
                SalesBinderAPI.DownloadBookImages(Args.Force);
            if(Args.SalesBinderContactsDownload)
                SalesBinderAPI.RetrieveAndSaveContacts();
            if(Args.SalesBinderAccountsDownload)
                SalesBinderAPI.RetrieveAndSaveAccounts();
            if(Args.SalesBinderInvoicesDownload)
                SalesBinderAPI.RetrieveAndSaveInvoices(topup_:!Args.Force);
            if (Args.SalesBinderInventoryList)
                salesBinderInventoryList(arguments, Args.Keys.SalesBinderInventoryListKey, false);
            if (Args.SalesBinderInventoryListCurrent)
                salesBinderInventoryList(arguments, Args.Keys.SalesBinderInventoryListCurrentKey, true);
            if (Args.KeepaLookupRefreshCurrentInventory)
                keepaLookupRefreshCurrentInventory();
            if (Args.KeepaLookupPrimeRecords)
                keepaLookupPrimeRecords();
            if (Args.GenerateSalesReport)
                generateSalesReport(arguments);
            if (Args.SalesBinderFindAndUploadImagesForInventoryWithoutAnImage)
                salesBinderFindAndUploadImagesForInventoryWithoutAnImage();
            if (Args.GenerateStockListFromInventory)
                generateStockListFromInventory(arguments);
            if (Args.SalesBinderReportNegativeQuantities)
                salesBinderReportNegativeQuantities();
            if (Args.SalesBinderInventoryUpdate)
                salesBinderInventoryUpdate(arguments);
            if (Args.ImagesForBarcodes)
                imagesForBarcodes(arguments);
            if (Args.Test)
                test();
        }

        private static void salesBinderAccountsList()
        {
            var list = SalesBinderAPI.Accounts;
                
            (list.Length==0 ? "No local accounts found" : list.FormatIntoColumnsReflectOnType(100))
                .ConsoleWriteLine();
        }
        
        private static void salesBinderInventoryList(Arguments arguments, string arg_, bool onlyCurrent_)
        {
            var targetPath = arguments.GetArgument(arg_);

            if (string.IsNullOrEmpty(targetPath))
            {
                "Where should the inventory list be saved to (fullpath to csv)?: "
                    .ConsoleWriteLine();
                targetPath = Console.ReadLine();
            }

            var list = SalesBinderAPI.RetrieveAndSaveInventory(true);

            if (onlyCurrent_)
                list = list.Where(x => x.Quantity > 0).ToArray();

            if (Args.KeepaAugmentList)
            {
                var excludeFromtree = new HashSet<string>(new[] {"Books", "Subjects"});
                var excludeBinding = new HashSet<string>(new[] {"Kindle Edition"});

                list.Where(x => !string.IsNullOrEmpty(x.BarCode))
                    .Select(x => (Item: x, Keepa: KeepaAPI.GetRecordForIdentifier(x.BarCode)))
                    .Where(x => x.Keepa != null)
                    .ForEach(l =>
                    {
                        if (string.IsNullOrEmpty(l.Item.Publisher))
                            l.Item.Publisher = l.Keepa.Manufacturer;

                        if (!string.IsNullOrEmpty(l.Keepa.Binding) && !excludeBinding.Contains(l.Keepa.Binding))
                            l.Item.Style = l.Keepa.Binding;

                        if (l.Keepa.CategoryTree != null)
                        {
                            var tree = l.Keepa.CategoryTree.Where(x => !excludeFromtree.Contains(x));

                            if (tree.Any())
                            {
                                l.Item.ProductType = string.Join(" / ", tree);
                                l.Item.ProductType2 = tree.Last();
                            }

                            l.Item.KidsOrAdult =
                                l.Keepa.CategoryTree.Any(x => x.ToLower().Contains("child"))
                                    ? "Kids"
                                    : "Adult";
                        }
                    });
            }

            using (var writer = new StreamWriter(targetPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(list
                    .OrderByDescending(x => Math.Abs(x.Quantity))
                    .ThenBy(x => x.Name));
            }
        }

        private static void keepaLookupRefreshCurrentInventory()
        {
            var items = SalesBinderAPI.Inventory.Where(x => x.Quantity > 0 && !string.IsNullOrEmpty(x.BarCode))
                .Select(x=>(Book:x,LastLookup:KeepaAPI.LastLookupTime(x.BarCode)))
                .Where(x=>x.LastLookup==null || (DateTime.Now-x.LastLookup.Value).TotalDays>1)
                .Select(x=>x.Book)
                .ToArray();
                
            // only want to do 100 as don't want to blow limits on keepa
            if(!items.Any())
                "Nothing to do - have been updated or attempted to be updated at least once in the last 24 hours"
                    .ConsoleWriteLine();
            else
            {
                var ids = items.Take(100).Select(x => x.BarCode).Distinct().ToArray();
                KeepaAPI.GetDetailsForIdentifiers(ids, forceRefresh_: true);
            }
        }

        private static void keepaLookupPrimeRecords()
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

        private static void generateListForWarehouse(Arguments arguments)
        {
            if (!arguments.HasArgument("gwl"))
            {
                "You need to supply an output filepath (which should be an xlsx) to write the list to".ConsoleWriteLine();
            }
            else if (!arguments.ArgumentDictionary.TryGetValue("gwl", out var filename) || string.IsNullOrEmpty(filename?.ToString()))
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
                    .WriteWarehouseFile(arguments.GetArgument("gwl"));
            }
        }
        
        private static void generateSalesReport(Arguments arguments)
        {
            string encaseStringWithComma(string input_) =>
                ((!string.IsNullOrEmpty(input_) && input_.Contains(','))
                 || (!string.IsNullOrEmpty(input_) && int.TryParse(input_, out var _)))
                    ? $"\"{input_}\""
                    : input_;

                var targetFile = arguments.GetArgument(Args.Keys.GenerateSalesReportKey);

                if (string.IsNullOrEmpty(targetFile))
                {
                    "Where should the salesreport (csv) be saved to (full path)? :".ConsoleWriteLine();
                    targetFile = Console.ReadLine();
                }

                if (SalesBinderAPI.Invoices == null)
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
                    File.WriteAllText(targetFile,sb.ToString());
                }
        }

        private static void topupKeepaRecords()
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
        
        private static void salesBinderFindAndUploadImagesForInventoryWithoutAnImage()
        {
            SalesBinderAPI.Inventory.Where(x => x.HasImageSaved() == false && x.Quantity > 0)
                .Select(x => (Item: x, Image: ExtensionMethods.FindImagePath(x.BarCode)))
                .ForEach(x =>
                {
                    $"{x.Item.BarCode} : found image : {x.Image}".ConsoleWriteLine();

                    if (!string.IsNullOrEmpty(x.Image))
                        SalesBinderAPI.UploadImage(x.Item.BarCode, x.Image);
                });
        }
        
        private static void generateStockListFromInventory(Arguments arguments)
        {
            var targetFile = arguments.GetArgument(Args.Keys.GenerateStockListFromInventoryKey);

            if (string.IsNullOrEmpty(targetFile))
            {
                "Where should the stocklist be saved to (full path to xlsx)? :".ConsoleWriteLine();
                targetFile = Console.ReadLine();
            }

            var inventory = SalesBinderAPI.RetrieveAndSaveInventory(true);


            var listOfFilters = new List<Func<SalesBinderInventoryItem, bool>>();
            listOfFilters.Add(x => x.Quantity > 0);

            {
                "Do you want to apply any filters? (y|n)?".ConsoleWrite();
                var a = Console.ReadKey();
                if (a.KeyChar == 'y')
                {
                    "".ConsoleWriteLine();
                    "Category (SB:KidsOrAdult) containing?:".ConsoleWrite();
                    var category = Console.ReadLine();
                    if (!string.IsNullOrEmpty(category))
                        listOfFilters.Add(f =>
                            !string.IsNullOrEmpty(f.KidsOrAdult) &&
                            f.ProductType.ToLower().Contains(category.ToLower()));
                    "Subcategory (SB:ProductType) containing?".ConsoleWrite();
                    var subcategory = Console.ReadLine();
                    if (!string.IsNullOrEmpty(subcategory))
                        listOfFilters.Add(f =>
                            !string.IsNullOrEmpty(f.ProductType) &&
                            f.ProductType.ToLower().Contains(subcategory.ToLower()));
                    "Publisher containing?:".ConsoleWrite();
                    var pub = Console.ReadLine();
                    if (!string.IsNullOrEmpty(pub))
                        listOfFilters.Add(f =>
                            !string.IsNullOrEmpty(f.Publisher) && f.Publisher.ToLower().Contains(pub.ToLower()));
                    "Territory restrictions (space separate multiple)?:".ConsoleWrite();
                    var tr = Console.ReadLine();
                    if (!string.IsNullOrEmpty(tr))
                    {
                        var sep = tr.ToLower().Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToList();
                        listOfFilters.Add(f => string.IsNullOrEmpty(f.SalesRestrictions) ||
                                               (f.SalesRestrictions.ToLower()
                                                   .Split(' ')
                                                   .Where(x => !string.IsNullOrEmpty(x))
                                                   .All(s => !sep.Contains(s))));
                    }
                    "Minimum quantity?:".ConsoleWrite();
                    var qty = Console.ReadLine();
                    if (!string.IsNullOrEmpty(qty))
                    {
                        if (int.TryParse(qty, out var qtyI))
                            listOfFilters.Add(f => f.Quantity >= qtyI);
                        else
                            "Could not parse given qty to an integer".ConsoleWriteLine();
                    }
                    "Maximum ckb net price?:".ConsoleWrite();
                    var maxPx = Console.ReadLine();
                    if (!string.IsNullOrEmpty(maxPx))
                    {
                        if (decimal.TryParse(maxPx, out var pxD))
                            listOfFilters.Add(f => f.Price <= pxD);
                        else
                            "Could not parse given price to a decimal".ConsoleWriteLine();
                    }
                    "Minimum ckb net price?:".ConsoleWrite();
                    var minPx = Console.ReadLine();
                    if (!string.IsNullOrEmpty(minPx))
                    {
                        if (decimal.TryParse(minPx, out var pxD))
                            listOfFilters.Add(f => f.Price >= pxD);
                        else
                            "Could not parse given price to a decimal".ConsoleWriteLine();
                    }
                }
            }


            if (!inventory.Any())
            {
                $"There is no salesbinder inventory locally that matches the filters.  run 'CKB.exe --sbid' to download inventory before trying to generate a stock list"
                    .ConsoleWriteLine();
            }
            else
            {
                "".ConsoleWriteLine();
                
                var filtered = inventory.Where(x => listOfFilters.All(f => f(x)))
                    .ToList();

                SalesBinderAPI.DownloadImagesForItems(filtered);

                filtered
                    .Select(x => (BarCode: x.BarCode, Image: ExtensionMethods.FindImagePath(x.BarCode), Item: x))
                    .OrderByDescending(x => x.Item, new StockListOrderer())
                    .WriteStockListFileInStandardGroups(targetFile);
            }
        }

        private static void salesBinderReportNegativeQuantities()
        {
            var negs = SalesBinderAPI.RetrieveAndSaveInventory(true)
                .Where(x => x.Quantity < 0);
                
            if(!negs.Any())
                "No negative quantities".ConsoleWriteLine();
            else
            {
                negs.Select(x=>new []{$"{x.Name}",$"{x.BarCode}",$"{x.Quantity}"})
                    .FormatIntoColumns(new[] {"Name","Barcode","Quantity"})
                    .ConsoleWriteLine();
                    
                Console.Write("Type 'yes' to zero out these negatives in salesbinder:");

                var isYes = Console.ReadLine();

                if (string.Compare("yes", isYes, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    var updates = negs.Select(n =>
                    {
                        var clone = n.CreateCloneFromProperties();
                        clone.Quantity = 0;
                        return clone;
                    });
                        
                    updates.ForEach(r=>r.DetectChanges(negs, true, true));
                }
            }
        }
        
        private static void salesBinderInventoryUpdate(Arguments arguments)
        {
                var sourcePath = arguments.GetArgument(Args.Keys.SalesBinderInventoryUpdateKey);
                bool doIt = true;

                while (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                {
                    "Path to csv file with updates (full path)?:"
                        .ConsoleWriteLine();

                    sourcePath = Console.ReadLine();
                    
                    if("exist".Equals(sourcePath) || "quit".Equals(sourcePath))
                        Environment.Exit(0);
                }

                var currentInventory = SalesBinderAPI.RetrieveAndSaveInventory(true);

                SalesBinderInventoryItem[] recs;

                using (var reader = new StreamReader(sourcePath))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    recs = csv.GetRecords<SalesBinderInventoryItem>().ToArray();
                }

                // check to make sure no barcodes or SKUs have pluses in them
                if (recs.SelectMany(r => new[] {r.BarCode, r.SKU})
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Any(x => x.Contains('+')))
                {
                    "Some of the SKUs or Barcodes in the file have a '+' in them.  Did Excel format screw up the formatting?"
                        .ConsoleWriteLine();
                    doIt = false;
                }

                if (doIt)
                {
                    var toUpdate = recs.Where(potentialUpdate =>
                            potentialUpdate.DetectChanges(currentInventory, Args.UpdateQuantities, false))
                        .ToArray();

                    if (!toUpdate.Any())
                        "No changes found.".ConsoleWriteLine();
                    else if (Args.Force)
                    {
                        var total = toUpdate.Length;
                        $"There are {total} records that will be updated.  ".ConsoleWrite();
                        Console.Write("Enter 'yes' to make the changes: ");
                        var entered = Console.ReadLine();
                        if (string.Compare("yes", entered.Trim(), StringComparison.OrdinalIgnoreCase) == 0)
                            toUpdate.ForEach((r,i) =>
                            {
                                $"{i+1}/{total}: ".ConsoleWrite();
                                r.DetectChanges(currentInventory, Args.UpdateQuantities, true);
                            });
                        else
                            "Changed aborted.".ConsoleWriteLine();
                    }
                }
        }
        
        private static void salesBinderCreateInventory(Arguments arguments)
        {
            if (!arguments.HasArgument("sbci") || !File.Exists(arguments.GetArgument("sbci")))
            {
                "You need to supply a valid path to a file in which the new inventory listed"
                    .ConsoleWriteLine();
            }
            else
            {
                var lines = File.ReadLines(arguments.GetArgument("sbci"));

                lines.Select(x=>x.Trim())
                    .GroupBy(l=>SalesBinderAPI.InventoryByBarcode.TryGetValue(l,out var _))
                    .OrderBy(x=>x.Key)
                    .Reverse()
                    .ForEach(x =>
                    {
                        // exists already if true
                        if (x.Key)
                        {
                            x.ForEach(l =>
                            {
                                $"Product with barcode {l} already exists (Name={SalesBinderAPI.InventoryByBarcode[l].Name})"
                                    .ConsoleWriteLine();
                            });
                        }
                        else
                        {
                            var recs = KeepaAPI.GetDetailsForIdentifiers(x.ToArray());
                    
                            x.ForEach(l =>
                            {
                                if (!recs.TryGetValue(l, out var rec))
                                    return;
                        
                                SalesBinderAPI.CreateInventory(l,rec);
                            });
                                
                        }
                    });
            }
            
        }
        
        private static void imagesForBarcodes(Arguments arguments)
        {
            var input = arguments.GetArgument(Args.Keys.ImagesForBarcodesKey);
            var output = arguments.GetArgument(Args.Keys.OutputToKey);

            while (string.IsNullOrEmpty(input) || !File.Exists(input))
            {
                "Enter source csv file location: ".ConsoleWrite();
                input = Console.ReadLine();
            }

            while (string.IsNullOrEmpty(output))
            {
                "Enter output xlsx file path: ".ConsoleWrite();
                output = Console.ReadLine();
            }

            var barcode = File.ReadAllLines(input);
                
            barcode.WriteBarCodeAndImageFile(output);
        }

        private static void test()
        {
            var c = SalesBinderAPI.RetrieveJsonForInventoryItem("5fff3b26-0644-4046-90fb-7bb93f71d8bf");
            c.ToString().ConsoleWriteLine();
        }
    }
}
