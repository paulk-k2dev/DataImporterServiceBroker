using SourceCode.SmartObjects.Client;

namespace DataImporter.Extensions
{
    public static class SmartPropertyCollectionExtensions
    {
        public static bool HasDateProperty(this SmartPropertyCollection properties)
        {
            foreach(SmartProperty property in properties)
            {
                switch (property.Type)
                {
                    case PropertyType.DateTime:
                    case PropertyType.Date:
                    case PropertyType.Time:
                        return true;
                }
            }

            return false;
        }
    }
}