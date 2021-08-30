using System;

namespace Keowi.XrmToolBox.Plugins.BulkDataFinder
{
    public class Search
    {
        public string InputData { get; set; }
        public bool IsFound { get; set; }
        public bool IsProcessed { get; set; }
        public string PrimaryAttribute { get; set; }
        public Guid RecordId { get; set; }
    }
}