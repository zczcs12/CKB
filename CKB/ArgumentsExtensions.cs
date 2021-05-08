using Utility.CommandLine;

namespace CKB
{
    public static class ArgumentsExtensions
    {
        public static bool HasArgument(this Arguments lookup, string key_) => lookup.ArgumentDictionary.ContainsKey(key_) &&
                                         !string.IsNullOrEmpty($"{lookup.ArgumentDictionary[key_]}");
        
        public static string GetArgument(this Arguments lookup, string key_) => lookup.HasArgument(key_) ? $"{lookup.ArgumentDictionary[key_]}" : null;

        
    }
}