﻿namespace Keowi.XrmToolBox.Plugins.BulkDataFinder
{
    public class SearchCriterias
    {
        public string Attribute { get; set; }
        public bool DisplayPrimaryAttribute { get; set; }
        public string Entity { get; set; }
        public string FetchXml { get; set; }
        public bool IgnoreHeader { get; set; }
        public string PrimaryAttribute { get; set; }
        public bool UseFilteredView { get; set; }
    }
}