using System;
using System.IO;
using System.Linq;
using System.Text;

namespace CKB
{
    internal static class EnvironmentSetup
    {
        internal const string SalesBinderApiKeyEnvName = "SALESBINDER_API_KEY"; 
        internal static readonly string SalesApiBinderKey =
            Environment.GetEnvironmentVariable(SalesBinderApiKeyEnvName, EnvironmentVariableTarget.User);


        internal const string KeepApiKeyEnvName = "KEEPA_API_KEY";
        internal static readonly string KeepaApiKey =
            Environment.GetEnvironmentVariable(KeepApiKeyEnvName, EnvironmentVariableTarget.User);

        internal const string DataDirectoryEnvName = "CKB_DATA_ROOT";
        internal static readonly string DataRoot =
            Environment.GetEnvironmentVariable(DataDirectoryEnvName, EnvironmentVariableTarget.User);

        public static bool IsProperlySetup(out string error_)
        {
            var setup = new (string EnvName, string Value, string Error)[]
            {
                (SalesBinderApiKeyEnvName, SalesApiBinderKey,
                    "Needs to be set to your SalesBinderAPI api key code so CKB.exe has access to the account"),
                (KeepApiKeyEnvName, KeepaApiKey,
                    "Needs to be set to your KeepAPI key code so CKB.exe has access to the account"),
                (DataDirectoryEnvName, DataRoot,
                    "Needs to be set to the directory in which records should be saved locally.  Ideally somehwere in your personal docs directory (c:\\users\\<you>\\...")
            };

            var errored = setup.Where(x => string.IsNullOrEmpty(x.Value));

            if (errored.Any())
            {
                var sb = new StringBuilder("Some environment variables need to be set before CKB.exe can run.");
                sb.AppendLine().AppendLine(errored.Select(x => new[] {x.EnvName, x.Error})
                    .FormatIntoColumns(new[] {"EnvVariable", "Reason it's needed"}));

                error_ = sb.ToString();

                return false;
            }
            
            // check to ensure data directory exists
            if (!Directory.Exists(DataRoot))
            {
                error_ =
                    $"Environment variable '{DataDirectoryEnvName}' is set to '{DataRoot}' but that directory does not exist";
                return false;
            }
            
            // check the subdirectories for Keepa/SalesBinder/have been created
            {
                var dirs = KeepaAPI.directories()
                    .Concat(SalesBinderAPI.directories())
                    .Concat(AbeBooksAPI.directories())
                    .Concat(BlackwellsAPI.directories())
                    .Select(x =>
                    {
                        if (Directory.Exists(x))
                            return (Exists: true, Path: x);

                        try
                        {
                            Directory.CreateDirectory(x);
                            return (Exists: true, Path: x);
                        }
                        catch
                        {
                            return (Exists: false, Path: x);
                        }
                    })
                    .Where(x => x.Exists == false)
                    .ToList();

                if (dirs.Any())
                {
                    var sb = new StringBuilder(
                        $"There was a problem creating the directories to store data.  The following could not be created:");
                    dirs.ForEach(x=>sb.AppendLine(x.Path));
                    error_ = sb.ToString();
                    return false;
                }
            }


            error_ = null;
            return true;

        }
    }
}