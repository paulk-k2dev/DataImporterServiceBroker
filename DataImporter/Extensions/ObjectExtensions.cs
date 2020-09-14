using System;
using System.ComponentModel;

namespace DataImporter.Extensions
{
    /// <summary>
    /// A set of extension methods for working with objects.
    /// </summary>
    public static class ObjectExtensions
    {
        /// <summary>
        /// Converts an object from one type to another 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static T ConvertTo<T>(this object value, T defaultValue = default(T))
        {
            // Deal with null
            if (value == null) return defaultValue;

            // Deal with same type conversions
            if (typeof(T) == value.GetType()) return (T)value;

            // Deal with converting to nullable of same type
            var inputType = value.GetType();
            if (inputType.IsValueType)
            {
                var nullableOfInputType = typeof(Nullable<>).MakeGenericType(inputType);
                if (typeof(T) == nullableOfInputType)
                {
                    var nullableWrapper = Activator.CreateInstance(nullableOfInputType, value);
                    return (T)nullableWrapper;
                }
            }

            // See if we can find a type converter...
            var converter = TypeDescriptor.GetConverter(typeof(T));
            if (!converter.CanConvertFrom(value.GetType()))
            {
                // Invalid conversion
                return defaultValue;
            }

            if (!converter.IsValid(value)) return defaultValue;

            return (T)converter.ConvertFrom(value);
        }        
    }
}