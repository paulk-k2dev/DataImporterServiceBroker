using System.Collections.Generic;

namespace DataImporter.Entities
{
    internal class ColumnDefinition
    {
        public string Name { get; set; }

        public List<string> CellReferences { get; set; }

        public ColumnDefinition()
        {
            CellReferences = new List<string>();
        }
    }
}