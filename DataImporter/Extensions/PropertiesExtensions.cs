using System;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

namespace DataImporter.Extensions
{
    public static class PropertiesExtensions
    {        
        /// <summary>
        /// Reads a <see cref="Property">Property</see> from the collection of Properties
        /// </summary>        
        /// <param name="properties">The collection of Properties</param>
        /// <param name="name">The name of the property to return</param>        
        /// <returns>The value of the property as a string if found, string.Empty if not</returns>
        public static string SafeRead(this Properties properties, string name)
        {
            return SafeRead(properties, name, string.Empty);
        }

        /// <summary>
        /// Reads a <see cref="Property">Property</see> from the collection of Properties
        /// </summary>
        /// <typeparam name="T">The type of property to read as</typeparam>
        /// <param name="properties">The collection of Properties</param>
        /// <param name="name">The name of the property to return</param>
        /// <param name="defaultValue">The default value to return if value not found</param>
        /// <returns>The value of the property, if found, the defaultValue if not</returns>
        public static T SafeRead<T>(this Properties properties, string name, T defaultValue)
        {
            if (!Contains(properties, name)) return defaultValue;

            var value = properties[name].Value;

            return value == null ? defaultValue : value.ConvertTo(defaultValue);
        }
        
        public static void Create(this Properties properties, string name, string displayName = "")
        {
            var meta = new MetaData
            {
                DisplayName = displayName,
                Description = ""
            };

            Create(properties, name, meta);
        }

        public static void Create(this Properties properties, string name, SoType soType, string displayName = "")
        {
            var meta = new MetaData
            {
                DisplayName = displayName,
                Description = ""
            };

            Create(properties, name, soType, meta);
        }

        private static void Create(this Properties properties, string name, MetaData metaData)
        {
            Create(properties, name, SoType.Text, metaData);
        }

        private static void Create(this Properties properties, string name, SoType soType, MetaData metaData = null)
        {
            if (properties.Contains(name)) return;

            if (metaData == null)
                metaData = new MetaData
                {
                    DisplayName = name,
                    Description = ""
                };

            properties.Create(new Property
            {
                Name = name,
                Type = GetType(soType),
                SoType = soType,
                MetaData = metaData,
                Value = null
            });
        }

        /// <summary>
        /// Boolean indicator to evaluate if the <see cref="Properties">Properties</see>
        /// collection contains a named property
        /// </summary>
        /// <param name="properties">The collection of properties</param>
        /// <param name="name">The name of the property to search</param>
        /// <returns>true if found, false if not</returns>
        private static bool Contains(this Properties properties, string name)
        {
            return properties.IndexOf(name) != -1 && !string.IsNullOrEmpty(name);
        }

        public static string GetType(SoType type)
        {
            switch (type)
            {
                case SoType.AutoGuid:
                case SoType.Guid:
                    return typeof(Guid).ToString();
                case SoType.Autonumber:
                case SoType.Number:
                    return typeof(int).ToString();
                case SoType.Decimal:
                    return typeof(decimal).ToString();
                case SoType.Default:
                case SoType.Memo:
                case SoType.Text:
                    return typeof(string).ToString();
                case SoType.Date:
                case SoType.DateTime:
                    return typeof(DateTime).ToString();
                case SoType.Time:
                    return typeof(TimeSpan).ToString();
                case SoType.YesNo:
                    return typeof(bool).ToString();
                case SoType.File:
                    return typeof(string).ToString();
                case SoType.HyperLink:
                case SoType.Image:
                case SoType.MultiValue:
                case SoType.Xml:
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), $"Unknown mapping for {type}");
            }
        }
    }
}