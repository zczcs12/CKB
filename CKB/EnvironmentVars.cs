using System;
using System.Linq;
using System.Text;

namespace CKB
{
    internal class EnvironmentVars
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

            if (!errored.Any())
            {
                error_ = null;
                return true;
            }

            var sb = new StringBuilder("Some environment variables need to be set before CKB.exe can run.");
            sb.AppendLine().AppendLine(errored.Select(x => new[] {x.EnvName, x.Error})
                .FormatIntoColumns(new[] {"EnvVariable", "Reason it's needed"}));

            error_ = sb.ToString();
            return false;

        }
    }
}