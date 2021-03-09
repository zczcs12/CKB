using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CKB
{
    public static class JObjectExtensions
    {
        public static bool TryExtract<T>(this JToken t_, string tag_, Func<JToken, (bool, T)> parser_, out T out_)
        {
            JToken t = null;
            try
            {
                t = t_?[tag_];
            }
            catch
            {
                t = null;
            }

            if (t == null)
            {
                out_ = default;
                return false;
            }

            var parseResult = parser_(t);
            out_ = parseResult.Item2;
            return parseResult.Item1;
        }

        public static bool TryExtractToken(this JToken t_, string tag_, out JToken t)
        {
            try
            {
                t = t_?[tag_];
            }
            catch
            {
                t = null;
            }

            return t != null;
        }

        public static JToken ExtractToken(this JToken t_, string tag_)
        {
            t_.TryExtractToken(tag_, out var ret);
            return ret;
        }

        public static bool TryExtractArray(this JToken t_, string tag_, out JArray t)
        {
            try
            {
                t=t_[tag_] as JArray;
            }
            catch
            {
                t = null;
            }

            return t != null;
        }

        public static JArray ExtractArray(this JToken t_, string tag_)
        {
            t_.TryExtractArray(tag_, out var t);
            return t;
        }
        
        public static bool TryExtractLong(this JToken t_, string tag_, out long out_)
            => t_.TryExtract(tag_, j =>
            {
                var results = long.TryParse(j.ToString(), out var l);
                return (results, l);
            }, out out_);

        public static long ExtractLong(this JToken t_, string tag_)
        {
            t_.TryExtractLong(tag_, out var ret);
            return ret;
        }

        public static bool TryExtractDecimal(this JToken t_, string tag_, out decimal out_)
            => t_.TryExtract(tag_, j =>
            {
                var results = decimal.TryParse(j.ToString(), out var l);
                return (results, l);
            }, out out_);

        public static decimal ExtractDecimal(this JToken t_, string tag_)
        {
            t_.TryExtractDecimal(tag_, out var ret);
            return ret;
        }

        public static bool TryExtractDouble(this JToken t_, string tag_, out double out_)
            => t_.TryExtract(tag_, j =>
            {
                var results = double.TryParse(j.ToString(), out var l);
                return (results, l);
            }, out out_);

        public static double ExtractDouble(this JToken t_, string tag_)
        {
            t_.TryExtractDouble(tag_, out var ret);
            return ret;
        }

        public static bool TryExtractInt(this JToken t_, string tag_, out int out_)
            => t_.TryExtract(tag_, j =>
            {
                var results = int.TryParse(j.ToString(), out var l);
                return (results, l);
            }, out out_);

        public static int ExtractInt(this JToken t_, string tag_)
        {
            t_.TryExtractInt(tag_, out var ret);
            return ret;
        }

        public static bool TryExtractString(this JToken t_, string tag_, out string out_)
            => t_.TryExtract(tag_, j => (true, j.ToString()), out out_);

        public static string ExtractString(this JToken t_, string tag_)
        {
            t_.TryExtractString(tag_, out var ret);
            return ret;
        }

        public static bool TryExtractDateExact(this JToken t_, string tag_, string format_, out DateTime out_)
            => t_.TryExtract(tag_, j =>
            {
                var results = DateTime.TryParseExact(j.ToString(), format_, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var l);
                return (results, l);
            }, out out_);

        public static DateTime ExtractDateExact(this JToken t_, string tag_, string format_)
        {
            t_.TryExtractDateExact(tag_, format_, out var ret);
            return ret;
        }

        public static bool TryExtractDate(this JToken t_, string tag_, out DateTime out_)
            => t_.TryExtract(tag_, j =>
            {
                var results = DateTime.TryParse(j.ToString(), CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var l);
                return (results, l);
            }, out out_);

        public static DateTime ExtractDate(this JToken t_, string tag_)
        {
            t_.TryExtractDate(tag_, out var ret);
            return ret;
        }

        private static readonly JsonSerializerSettings Settings =
            new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
            };

        public static string ToJson(this object data_)
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(data_, Settings);
        }

        public static T Deserialize<T>(this string json_)
        {
            return Newtonsoft.Json.JsonConvert.DeserializeObject<T>(json_);
        }

        public static string FormatIntoColumns(this IEnumerable<string[]> given_, int? maxColumnWidth_=null)
        {
            var numCols = given_.First().Length;

            var widths = Enumerable.Range(0, numCols)
                .Select(index => given_.Select(g => string.IsNullOrEmpty(g[index]) ? 0 : g[index].Length))
                .Select(x => x.Max())
                .Select(x=>maxColumnWidth_.HasValue ? Math.Min(maxColumnWidth_.Value,x) : x)
                .ToArray();

            return string.Join(Environment.NewLine,
                given_.Select(arr => string.Join(" | ",
                    arr.Select((s, index) =>
                        s == null 
                            ? string.Empty.PadRight(widths[index])
                            : s.Length>widths[index]
                            ? trimIfLonger_(s,widths[index])
                            : s.PadRight(widths[index])))));
        }

        public static string FormatIntoColumns(this IEnumerable<string[]> given_, string[] headings_, int? maxColumnWidth_=null)
            => new List<string[]> {headings_}.Concat(given_).FormatIntoColumns(maxColumnWidth_);

        private static string trimIfLonger_(this string str_, int? maxLength_)
            => maxLength_ == null || str_.Length < maxLength_ ? str_ : $"{str_.Substring(0, maxLength_.Value - 3)}...";
        
        public static string FormatIntoMarkdownTable(this IEnumerable<string[]> given_, string[] headings_)
        {
            var numCols = headings_.Length;

            string toLine(string[] cells_)
            {
                var contents = string.Join("|", cells_);
                return $"|{contents}|";
            }

            return string.Join(Environment.NewLine, new[]
                {
                    headings_,
                    "---".CreateArrayRep(numCols)
                }.Concat(given_)
                .Select(toLine));
        }

        public static T[] CreateArrayRep<T>(this T value_, int length_)
        {
            T[] ret = new T[length_];
            for (int i = 0; i < length_; ++i)
                ret[i] = value_;
            return ret;
        }

        public static string FormatIntoColumnsReflectOnType<T>(this IEnumerable<T> items_, int? maxColumnWidth_=null) where T : class
        {
            var props = typeof(T) == typeof(string) || typeof(T).IsValueType
                ? null
                : typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.GetProperty | BindingFlags.Public)
                    .ToArray();

            var formats = props.ToDictionary(p => p.Name,
                p => null as string); // asBoundToGrid_ == false || !p.TryGetSingleAttribute<DisplayFormatAttribute>(out var att) ? null : att.StringFormat);

            var list = new List<string[]> {props.Select(x => x.Name).ToArray()};

            items_.ForEach(i =>
                list.Add(props.Select(p =>
                {
                    var o = p.GetValue(i);

                    return o == null
                        ? null
                        : o is IFormattable formattable
                            ? formattable.ToString(formats[p.Name], CultureInfo.InvariantCulture)
                            : $"{o}";
                }).ToArray())
            );

            return list.FormatIntoColumns(maxColumnWidth_);
        }

    }
}