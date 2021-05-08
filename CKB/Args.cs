using Utility.CommandLine;

namespace CKB
{
    public class Args
    {
        [Argument('a', Keys.SalesBinderAccountsDownloadKey, "Sales Binder Account Download")]
        public static bool SalesBinderAccountsDownload { get; set; }

        [Argument('c', Keys.SalesBinderInventoryDownloadKey, "Sales Binder Inventory Download")]
        public static bool SalesBinderInventoryDownload { get; set; }

        [Argument('d', Keys.SalesBinderInventoryListKey, "Sales Binder Inventory List => csv")]
        public static bool SalesBinderInventoryList { get; set; }

        [Argument('e', Keys.SalesBinderInventoryListCurrentKey, "Sales Binder Inventory List (Current) => csv")]
        public static bool SalesBinderInventoryListCurrent { get; set; }

        [Argument('f', Keys.SalesBinderImageDownloadKey, "Sales Binder Image Download")]
        public static bool SalesBinderImageDownload { get; set; }

        [Argument('g', Keys.SalesBinderContactsDownloadKey, "Sales Binder Contact Download")]
        public static bool SalesBinderContactsDownload { get; set; }

        [Argument('i', Keys.SalesBinderInvoicesDownloadKey, "Sales Binder Invoices Download")]
        public static bool SalesBinderInvoicesDownload { get; set; }

        [Argument('j', Keys.SalesBinderInventoryUpdateKey, "Sales Binder Inventory Update (from file)")]
        public static bool SalesBinderInventoryUpdate { get; set; }

        [Argument('u', Keys.SalesBinderFindAndUploadImagesForInventoryWithoutAnImageKey, "Sales Binder Image Upload (find for inventory without one)")]
        public static bool SalesBinderFindAndUploadImagesForInventoryWithoutAnImage { get; set; }

        [Argument('k', Keys.ForceKey, "Force updates")]
        public static bool Force { get; set; }

        [Argument('/', Keys.KeepaLookupPrimeRecordsKey, "Keepa lookup prime records.  Try to get a keepa record for anthing we've not tried before.")]
        public static bool KeepaLookupPrimeRecords { get; set; }

        [Argument('l', Keys.KeepaLookupRefreshCurrentInventoryKey, "Keepa lookup refresh inventory - if not updated in the last 24 hours")]
        public static bool KeepaLookupRefreshCurrentInventory { get; set; }

        [Argument(')', Keys.KeepaAugmentListKey, "Keepa augment (the salesbinder csv list)")]
        public static bool KeepaAugmentList { get; set; }

        [Argument('n', Keys.GenerateSalesReportKey, "Generate sales report csv")]
        public static bool GenerateSalesReport { get; set; }

        [Argument('t', Keys.TestKey, "Test code")] 
        public static bool Test { get; set; }

        [Argument('#', Keys.GenerateStockListFromInventoryKey, "Generate stock list from inventory => xlsx")]
        public static bool GenerateStockListFromInventory { get; set; }

        [Argument('*', Keys.SalesBinderReportNegativeQuantitiesKey, "Sales Binder report negative quantities")]
        public static bool SalesBinderReportNegativeQuantities { get; set; }

        [Argument('&', Keys.UpdateQuantitiesKey, "Include changing quantities of SalesBinder updates")]
        public static bool UpdateQuantities { get; set; }

        [Argument(',', Keys.ImagesForBarcodesKey, "Csv of barcodes => xlsx of csv/images")]
        public static bool ImagesForBarcodes { get; set; }

        [Argument('-', Keys.OutputToKey, "Output file to (only applies to some options)")]
        public static bool OutputTo { get; set; }

        public static class Keys
        {
            public const string SalesBinderAccountsDownloadKey = "sbad";
            public const string SalesBinderInventoryDownloadKey = "sbid";
            public const string SalesBinderInventoryListKey = "sbil";
            public const string SalesBinderInventoryListCurrentKey = "sbilc";
            public const string SalesBinderImageDownloadKey = "sbimd";
            public const string SalesBinderContactsDownloadKey = "sbcd";
            public const string SalesBinderInvoicesDownloadKey = "sbinvd";
            public const string SalesBinderInventoryUpdateKey = "sbiu";
            public const string SalesBinderFindAndUploadImagesForInventoryWithoutAnImageKey = "sbimu";
            public const string ForceKey = "force";
            public const string KeepaLookupPrimeRecordsKey = "klp";
            public const string KeepaLookupRefreshCurrentInventoryKey = "klri";
            public const string KeepaAugmentListKey = "ka";
            public const string GenerateSalesReportKey = "gsr";
            public const string TestKey = "test";
            public const string GenerateStockListFromInventoryKey = "gsli";
            public const string SalesBinderReportNegativeQuantitiesKey = "sbnq";
            public const string UpdateQuantitiesKey = "updatequantities";
            public const string ImagesForBarcodesKey = "gib";
            public const string OutputToKey = "output";
        }
    }
}