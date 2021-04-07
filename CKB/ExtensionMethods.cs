using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CKB
{
    public static class ExtensionMethods
    {
        public static void ForEach<T>(this IEnumerable<T> items_, Action<T> onEach_)
        {
            foreach (var i in items_)
                onEach_(i);
        }

        public static string ImageFilePath(this SalesBinderInventoryItem inventoryItem, ImageSize size_)
        {
            switch (size_)
            {
                case ImageSize.Small:
                    return $@"{SalesBinderAPI.SAVE_PATH_ROOT}\images\{inventoryItem.SKU}_small.jpg";
                case ImageSize.Medium:
                    return $@"{SalesBinderAPI.SAVE_PATH_ROOT}\images\{inventoryItem.SKU}_medium.jpg";
                case ImageSize.Large:
                    return $@"{SalesBinderAPI.SAVE_PATH_ROOT}\images\{inventoryItem.SKU}_medium.jpg";
                default:
                    return null;
            }
        }

        public static bool HasImageSaved(this SalesBinderInventoryItem inventoryItem)
            => Enum.GetValues(typeof(ImageSize)).Cast<ImageSize>().Select(iss => inventoryItem.ImageFilePath(iss))
                .Any(File.Exists);
        
        public static void ConsoleWriteLine(this string str_) => Console.WriteLine(str_);

        public static void ConsoleWrite(this string str_) => Console.Write(str_);

        public static string Join(this string[] items_, string sep_) => string.Join(sep_, items_);

        public static void WriteWarehouseFile(this IEnumerable<(string Identifer, int Qty, KeepaRecord Record)> books_,
            string saveTo_)
        {
            var spreadSheetDocument = SpreadsheetDocument.Create(saveTo_, SpreadsheetDocumentType.Workbook);
            var workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            var workSheetPart = workbookpart.AddNewPart<WorksheetPart>();
            workSheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = spreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            var sheet = new Sheet
            {
                Id = spreadSheetDocument.WorkbookPart.GetIdOfPart(workSheetPart),
                SheetId = 1,
                Name = "Stock"
            };
            sheets.Append(sheet);

            // headings
            var settings =
                new (string ExcelColRef, string Heading,
                    Func<(string Identifer, int Qty, KeepaRecord Record), CellValue> ValueGetter, CellValues Type)
                    []
                    {
                        ("A", "Title", x => new CellValue(x.Record?.Title), CellValues.String),
                        ("B", "Quantity", x => new CellValue(Convert.ToDouble(x.Qty)), CellValues.Number),
                        ("C", "Author", x => new CellValue(x.Record?.Author), CellValues.Number),
                        ("D", "BarCode", x => new CellValue(x.Identifer), CellValues.String),
                        ("E", "EanList", x => new CellValue(x.Record?.EanList.Join(",")), CellValues.String),
                        ("E", "Binding", x => new CellValue(x.Record?.Binding), CellValues.String),
                        ("F", "ISBN10", x => new CellValue(x.Record?.Asin), CellValues.String),
                    };

            uint rowNumber = 1;

            settings.ForEach(s =>
            {
                var cell = workSheetPart.InsertCellInWorksheet(s.ExcelColRef, rowNumber);
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(s.Heading);
            });
            
            var sheetData = sheet.GetFirstChild<SheetData>();

            foreach (var book in books_)
            {
                rowNumber += 1;

                settings.ForEach(s =>
                {
                    var cell = workSheetPart.InsertCellInWorksheet(s.ExcelColRef, rowNumber);
                    cell.DataType = s.Type;
                    cell.CellValue = s.ValueGetter(book);
                });

                var row = workSheetPart.Worksheet.Descendants<Row>()
                    .FirstOrDefault(r => r.RowIndex == (uint) rowNumber);

                if (row != null)
                {
                    row.Height = 60;
                    row.CustomHeight = true;
                }
                
                var imagePath = book.Record?.ImagePaths()?[0];
                
                if (File.Exists(imagePath))
                {
                    workSheetPart.AddImage(imagePath, book.Identifer, settings.Count() + 1, (int) (rowNumber), ProcessImageForExcel);
                }
            }

            workbookpart.Workbook.Save();

            spreadSheetDocument.Close();

        }

        public static void WriteExcelFile(this IEnumerable<SalesBinderInventoryItem> books_, string saveTo_)
        {
            var spreadSheetDocument = SpreadsheetDocument.Create(saveTo_, SpreadsheetDocumentType.Workbook);
            var workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            var workSheetPart = workbookpart.AddNewPart<WorksheetPart>();
            workSheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = spreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            var sheet = new Sheet
            {
                Id = spreadSheetDocument.WorkbookPart.GetIdOfPart(workSheetPart),
                SheetId = 1,
                Name = "Stock"
            };
            sheets.Append(sheet);

            // headings
            var settings =
                new (string ExcelColRef, string Heading, Func<SalesBinderInventoryItem, CellValue> ValueGetter, CellValues Type)
                    []
                    {
                        ("A", "Barcode", x => new CellValue(x.Name), CellValues.String),
                        ("B", "Quantity", x => new CellValue(Convert.ToDouble(x.Quantity)), CellValues.Number),
                        ("C", "Price", x => new CellValue(x.Price), CellValues.Number),
                        ("D", "SKU", x => new CellValue(x.SKU), CellValues.String),
                        ("E", "BarCode", x => new CellValue(x.BarCode), CellValues.String),
                        ("F", "Style", x => new CellValue(x.Style), CellValues.String),
                        ("G", "KidsOrAdult", x => new CellValue(x.KidsOrAdult), CellValues.String),
                        ("H", "Publisher", x => new CellValue(x.Publisher), CellValues.String),
                    };

            uint rowNumber = 1;

            settings.ForEach(s =>
            {
                var cell = workSheetPart.InsertCellInWorksheet(s.ExcelColRef, rowNumber);
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(s.Heading);
            });

            foreach (var book in books_)
            {
                rowNumber += 1;

                settings.ForEach(s =>
                {
                    var cell = workSheetPart.InsertCellInWorksheet(s.ExcelColRef, rowNumber);
                    cell.DataType = s.Type;
                    cell.CellValue = s.ValueGetter(book);
                });

                var imagePath = book.ImageFilePath(ImageSize.Small);

                if (File.Exists(imagePath))
                {
                    workSheetPart.AddImage(imagePath, book.SKU, settings.Count() + 1, (int) (rowNumber),ProcessImageForExcel);
                }
                else
                {
                    var keepa = KeepaAPI.GetRecordForIdentifier(book.BarCode);
                    if (keepa != null)
                    {
                        var pathToImage = keepa.ImagePaths().FirstOrDefault(File.Exists);
                        if (!string.IsNullOrEmpty(pathToImage))
                        {
                            workSheetPart.AddImage(pathToImage, book.SKU, settings.Count() + 1, (int) (rowNumber),ProcessImageForExcel);
                        }
                    }
                }
                
                var row = workSheetPart.Worksheet.Descendants<Row>()
                    .FirstOrDefault(r => r.RowIndex == (uint) rowNumber);

                if (row != null)
                {
                    row.Height = 60;
                    row.CustomHeight = true;
                }

            }

            workbookpart.Workbook.Save();

            spreadSheetDocument.Close();

        }

        public static void WriteStockListFile(
            this IEnumerable<(string Barcode, string ImagePath, SalesBinderInventoryItem Book)> items_, string saveTo_)
        {
            var spreadSheetDocument = SpreadsheetDocument.Create(saveTo_, SpreadsheetDocumentType.Workbook);
            var workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            var workSheetPart = workbookpart.AddNewPart<WorksheetPart>();
            workSheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = spreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            var sheet = new Sheet
            {
                Id = spreadSheetDocument.WorkbookPart.GetIdOfPart(workSheetPart),
                SheetId = 1,
                Name = "Stock"
            };
            sheets.Append(sheet);

            // headings
            var settings =
                new (string ExcelColRef, string Heading,
                    Func<(string Barcode, string ImagePath, SalesBinderInventoryItem Book), CellValue> ValueGetter, CellValues
                    Type)
                    []
                    {
                        ("A", "Barcode", x => new CellValue(x.Barcode), CellValues.String),
                        ("B", "Image", x => new CellValue(string.Empty), CellValues.String),
                        ("C", "Product description", x => new CellValue(x.Book?.Name ?? string.Empty), CellValues.String),
                        ("D", "Type", x=> new CellValue(x.Book?.ProductType ?? string.Empty), CellValues.String),
                        ("E", "SubType", x=> new CellValue(x.Book?.ProductType2 ?? string.Empty), CellValues.String),
                        ("F", "SubType2", x=> new CellValue(x.Book?.ProductType3 ?? string.Empty), CellValues.String),
                        ("G", "Author", x => new CellValue(x.Book?.Author ?? string.Empty), CellValues.String),
                        ("H", "Full RRP", x => new CellValue(x.Book?.FullRRP ?? string.Empty), CellValues.Number),
                        ("I", "Style", x => new CellValue(x.Book?.Style ?? string.Empty), CellValues.String),
                        ("J", "Publisher", x => new CellValue(x.Book?.Publisher ?? string.Empty), CellValues.String),
                        ("K", "Clearance?", x => new CellValue(string.Empty), CellValues.String),
                        ("L", "Case size", x => new CellValue(x.Book?.PackSize ?? string.Empty), CellValues.String),
                        ("M", "VAT Rate", x => new CellValue(x.Book?.VAT ?? string.Empty), CellValues.String),
                        ("N", "Condition", x => new CellValue(x.Book?.Condition), CellValues.String),
                        ("O", "Territory Restrictions", x => new CellValue(x.Book?.SalesRestrictions), CellValues.String),
                        ("P", "CKB Net Price", x => new CellValue(x.Book == null ? string.Empty : $"£{x.Book.Price:0.00}"), CellValues.String),
                        ("Q", "Qty", x => new CellValue(x.Book == null ? string.Empty : $"{x.Book?.Quantity:#,###}"), CellValues.String),
                        ("R", "Order Quantity", x => new CellValue(string.Empty), CellValues.String),
                    };

            uint rowNumber = 1;

            settings.ForEach(s =>
            {
                var cell = workSheetPart.InsertCellInWorksheet(s.ExcelColRef, rowNumber);
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(s.Heading);
            });

            foreach (var book in items_)
            {
                rowNumber += 1;

                settings.ForEach(s =>
                {
                    var cell = workSheetPart.InsertCellInWorksheet(s.ExcelColRef, rowNumber);
                    cell.DataType = s.Type;
                    cell.CellValue = s.ValueGetter(book);
                });

                if (!string.IsNullOrEmpty(book.ImagePath) && File.Exists(book.ImagePath))
                {
                    var row = workSheetPart.Worksheet.Descendants<Row>()
                        .FirstOrDefault(r => r.RowIndex == (uint) rowNumber);

                    if (row != null)
                    {
                        row.Height = 60;
                        row.CustomHeight = true;
                    }

                    workSheetPart.AddImage(book.ImagePath, string.Empty, 2, (int) (rowNumber),ProcessImageForExcel);
                }
            }

            workbookpart.Workbook.Save();

            spreadSheetDocument.Close();
        }
        
        public static Bitmap ProcessImageForExcel(Bitmap img_)
        {
            try
            {
                return img_.ResizeImageToHeight(60);
            }
            catch
            {
                return img_;
            }
        }

        public static string FindImagePath(string bookIdentifier_)
        {
            if (SalesBinderAPI.InventoryByBarcode.TryGetValue(bookIdentifier_, out var book_))
            {
                var images = new[]
                {
                    ImageSize.Small,
                    ImageSize.Medium,
                    ImageSize.Large
                }.Select(size => book_.ImageFilePath(size));

                var found = images.FirstOrDefault(File.Exists);

                if (!string.IsNullOrEmpty(found))
                    return found;
            }

            // now try keepa
            {
                var record = KeepaAPI.HaveLocalRecord(bookIdentifier_)
                    ? KeepaAPI.GetRecordForIdentifier(bookIdentifier_)
                    : null;

                if (record != null)
                {
                    var paths = record.ImagePaths();

                    var found = paths == null ? paths.FirstOrDefault(File.Exists) : null;

                    if (!string.IsNullOrEmpty(found))
                        return found;
                }
            }
            
            // now try manually downloaded
            {
                var path = $@"{Environment.GetEnvironmentVariable("CKB_DATA_ROOT",EnvironmentVariableTarget.User)}\ManualImages\{bookIdentifier_}.jpg";
                if (File.Exists(path))
                    return path;
            }

            // now try abebooks
            {
                var found = AbeBooksAPI.TryGetImageForIdentifier(bookIdentifier_);

                if (File.Exists(found))
                    return found;
            }
            
            // not try blackwells
            {
                var found = BlackwellsAPI.TryGetImageForIdentifier(bookIdentifier_);

                if (File.Exists(found))
                    return found;
            }

            return null;
        }

        public static int ToInt(this IConvertible d) => Convert.ToInt32(d);

        public static int ToEpochTime(this DateTime date_) => (date_ - new DateTime(1970, 1, 1)).TotalSeconds.ToInt();

        public static string GetTextFromEmbeddedResource(string name_)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name_))
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd();
        }

        private static ConcurrentDictionary<string, PropertyInfo[]> _propsCache =
            new ConcurrentDictionary<string, PropertyInfo[]>();
        private static PropertyInfo[] GetProperties(Type t_)
        {
            if (_propsCache.TryGetValue(t_.FullName, out var ret))
                return ret;

            ret = t_.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty |
                                   BindingFlags.SetProperty);

            _propsCache.TryAdd(t_.FullName, ret);

            return ret;
        }
        
        public static T CreateCloneFromProperties<T>(this T i_) where T : new()
        {
            var props = GetProperties(typeof(T));

            var ret = new T();
            
            props.ForEach(p=>p.SetValue(ret,p.GetValue(i_)));

            return ret;
        }
    }
}