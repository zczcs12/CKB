using Utility.CommandLine;

namespace CKB
{
    public class Args
    {
        [Argument('a', Keys.SalesBinderAccountsDownload, "Sales Binder Account Download")]
        public static bool SalesBinderAccountsDownload { get; set; }

        [Argument('c', Keys.SalesBinderInventoryDownload, "Sales Binder Inventory Download")]
        public static bool SalesBinderInventoryDownload { get; set; }

        [Argument('d', Keys.SalesBinderInventoryList, "Sales Binder Inventory List => csv")]
        public static bool SalesBinderInventoryList { get; set; }

        [Argument('e', Keys.SalesBinderInventoryListCurrent, "Sales Binder Inventory List (Current) => csv")]
        public static bool SalesBinderInventoryListCurrent { get; set; }

        [Argument('f', Keys.SalesBinderImageDownload, "Sales Binder Image Download")]
        public static bool SalesBinderImageDownload { get; set; }

        [Argument('g', Keys.SalesBinderContactsDownload, "Sales Binder Contact Download")]
        public static bool SalesBinderContactsDownload { get; set; }

        [Argument('i', Keys.SalesBinderInvoicesDownload, "Sales Binder Invoices Download")]
        public static bool SalesBinderInvoicesDownload { get; set; }

        [Argument('j', Keys.SalesBinderInventoryUpdate, "Sales Binder Inventory Update (from file)")]
        public static bool SalesBinderInventoryUpdate { get; set; }

        [Argument('u', Keys.SalesBinderFindAndUploadImagesForInventoryWithoutAnImage, "Sales Binder Image Upload (find for inventory without one)")]
        public static bool SalesBinderFindAndUploadImagesForInventoryWithoutAnImage { get; set; }

        [Argument('k', Keys.Force, "Force updates")]
        public static bool Force { get; set; }

        [Argument('/', Keys.KeepaLookupPrimeRecords, "Keepa lookup prime records.  Try to get a keepa record for anthing we've not tried before.")]
        public static bool KeepaLookupPrimeRecords { get; set; }

        [Argument('l', Keys.KeepaLookupRefreshCurrentInventory, "Keepa lookup refresh inventory - if not updated in the last 24 hours")]
        public static bool KeepaLookupRefreshCurrentInventory { get; set; }

        [Argument(')', Keys.KeepaAugmentList, "Keepa augment (the salesbinder csv list)")]
        public static bool KeepaAugmentList { get; set; }

        [Argument('n', Keys.GenerateSalesReport, "Generate sales report csv")]
        public static bool GenerateSalesReport { get; set; }

        [Argument('t', Keys.Test, "Test code")] 
        public static bool Test { get; set; }

        [Argument('#', Keys.GenerateStockListFromInventory, "Generate stock list from inventory => xlsx")]
        public static bool GenerateStockListFromInventory { get; set; }

        [Argument('*', Keys.SalesBinderReportNegativeQuantities, "Sales Binder report negative quantities")]
        public static bool SalesBinderReportNegativeQuantities { get; set; }

        [Argument('&', Keys.UpdateQuantities, "Include changing quantities of SalesBinder updates")]
        public static bool UpdateQuantities { get; set; }

        [Argument(',', Keys.ImagesForBarcodes, "Csv of barcodes => xlsx of csv/images")]
        public static bool ImagesForBarcodes { get; set; }

        [Argument('-', Keys.OutputTo, "Output file to (only applies to some options)")]
        public static bool OutputTo { get; set; }

        public static class Keys
        {
            public const string SalesBinderAccountsDownload = "sbad";
            public const string SalesBinderInventoryDownload = "sbid";
            public const string SalesBinderInventoryList = "sbil";
            public const string SalesBinderInventoryListCurrent = "sbilc";
            public const string SalesBinderImageDownload = "sbimd";
            public const string SalesBinderContactsDownload = "sbcd";
            public const string SalesBinderInvoicesDownload = "sbinvd";
            public const string SalesBinderInventoryUpdate = "sbiu";
            public const string SalesBinderFindAndUploadImagesForInventoryWithoutAnImage = "sbimu";
            public const string Force = "force";
            public const string KeepaLookupPrimeRecords = "klp";
            public const string KeepaLookupRefreshCurrentInventory = "klri";
            public const string KeepaAugmentList = "ka";
            public const string GenerateSalesReport = "gsr";
            public const string Test = "test";
            public const string GenerateStockListFromInventory = "gsli";
            public const string SalesBinderReportNegativeQuantities = "sbnq";
            public const string UpdateQuantities = "updatequantities";
            public const string ImagesForBarcodes = "gib";
            public const string OutputTo = "output";
        }
    }
}