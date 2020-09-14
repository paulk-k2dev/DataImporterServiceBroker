using System;
using SourceCode.SmartObjects.Client;

namespace DataImporter.Extensions
{
    public static class SmartPropertyExtensions
    {
        private const string DateFormat = "yyyy-MM-dd";
        private const string TimeFormat = "HH:mm:ss";

        public static string GetValue(this SmartProperty property, object value)
        {
            var stringValue = value.ToString();

            if (string.IsNullOrEmpty(stringValue)) return string.Empty;

            switch (property.Type)
            {
                case PropertyType.Date:
                    // try default date conversion
                    try
                    {
                        var date = DateTime.Parse(stringValue);
                        return date.ToString(DateFormat);
                    }
                    // if fails then probably an excel date
                    catch
                    {
                        var date = DateTime.FromOADate(Convert.ToDouble(value.ToString()));
                        return date.ToString(DateFormat);
                    }
                case PropertyType.DateTime:
                    // try default date conversion
                    try
                    {
                        var dateTime = DateTime.Parse(stringValue);
                        return dateTime.ToString($"{DateFormat} {TimeFormat}");
                    }
                    // if fails then probably an excel date
                    catch
                    {
                        var dateTime = DateTime.FromOADate(Convert.ToDouble(value.ToString()));
                        return dateTime.ToString($"{DateFormat} {TimeFormat}");
                    }
                case PropertyType.Time:
                    // try default date conversion
                    try
                    {
                        var time = DateTime.Parse(stringValue);
                        return time.ToString(TimeFormat);
                    }
                    // if fails then probably an excel date
                    catch
                    {
                        var time = DateTime.FromOADate(Convert.ToDouble(value.ToString()));
                        return time.ToString(TimeFormat);
                    }
                default:
                    return value.ToString();
            }
        }
    }
}