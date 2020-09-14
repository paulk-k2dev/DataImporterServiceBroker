using System.Collections.Generic;

namespace DataImporter.Entities
{
    internal class ColumnDefinition
    {
        public string Name { get; set; }
        public List<int> Positions { get; set; }

        public ColumnDefinition()
        {
            Positions = new List<int>();
        }
    }
}