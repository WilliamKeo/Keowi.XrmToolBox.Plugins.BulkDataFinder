using Keowi.XrmToolBox.Plugins.BulkDataFinder.AppCode;
using System.Collections.Generic;

namespace Keowi.XrmToolBox.Plugins.BulkDataFinder
{
    public class SearchCriterias
    {
        public string Attribute { get; set; }
        public List<string> Columns { get; set; }
        public ColumnsOption ColumnsOption { get; set; }
        public string Entity { get; set; }
        public string FetchXml { get; set; }
        public bool IgnoreHeader { get; set; }
        public string PrimaryAttribute { get; set; }
        public bool UseFilteredView { get; set; }
    }
}