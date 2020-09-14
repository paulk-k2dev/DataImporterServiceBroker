using System;
using System.Text.RegularExpressions;

namespace DataImporter.Extensions
{
    public static class StringExtensions
    {
        /// <summary>
        /// Formats the string in such a way that it is suitable for referring to as a SmartObject
        /// Property name.
        /// </summary>
        /// <param name="name">The column name to format</param>
        /// <param name="mode">The mode to use to handle space character replacements</param>
        /// <returns>The formatted string</returns>
        public static string FormatColumnName(this string name, string mode)
        {
            if (string.IsNullOrWhiteSpace(name)) return string.Empty;

            var result = mode.Equals("Replace", StringComparison.InvariantCultureIgnoreCase)
                ? name.Replace(" ", "_")
                : name.Replace(" ", "");

            // remove any CRs / LFs
            // also get rid of other (non-underscore or hyphen) punctuation
            // which is also invalid for a SmartObject Property Name
            result = Regex.Replace(result, @"[\r\n\p{P}\p{S}-[-_]]", "");

            return result;
        }
    }
}