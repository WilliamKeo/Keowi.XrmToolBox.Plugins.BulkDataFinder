using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Keowi.XrmToolBox.Plugins.BulkDataFinder
{
    // Do not forget to update version number and author (company attribute) in AssemblyInfo.cs class
    // To generate Base64 string for Images below, you can use https://www.base64-image.de/
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "Bulk Data Finder"),
        ExportMetadata("Description", "Use input file to cross search into Dynamics 365 or Dataverse and get the results on a Excel file."),
        // Please specify the base64 content of a 32x32 pixels image
        ExportMetadata("SmallImageBase64", "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAFQ0lEQVRYR+1Xa0xTZxh+TksRaYFSQOsqbIhB2VjQZTHZTMxMXLK4LbsJPzSaoGbVtVhaoIigq5jSABUBHSNBrYkomSYTsqXZkiUMfi2SETRbRIZWmOUyXEsvsNKWc5bz0RZL23FxiX/8frQ9b9/L8z3v5fsOBY2GwXNc1AsALxh47gxQC7rg/2wJHocDUBSop7qM/KbmJOwnNTg4SGImJiaCoih81dGBk0NDz9yYr09NobO0FFwuFzExMYiKigry6X+mGIYJ2vSQ2Yz05mYwCwyWi6h9+3Z8tGtXRDN/WGq9UsmwlAg4HNzRasHj8RBTVAR3QsJyYwb0OS4XZnU68vzmiRMwu1ygATAUBXa37LeXpsn3/CBiGPy+dy9ezczEW2Vl+CUmZsUAdszMoKuqCiPj45A0N/+nn6BJWJacjCqZDHVtbSgaGFgZAJrGr7m5eCM7G/t1OrS63UsH8LLFgkcNDbA5HBDW1YU1/ATAZqEQOosFYKt8weLb7XDo9ZilafBVKrhFoqUDoNxu0FotMVhVUgK3QBBkTLlc8Gq14HA4SJfL8SglJcR5YVISzsnl6Lp9G+8YjYGWi4Qi5DAy5efjlbQ0ZBYX44+4uCC7j71e3DpzhsgM7e04eOdOiF9rYSGECQnIUShwd5Hdz82CBYNIl5qK4wcPov7GDSjv3ZsPwDDo27MHOdnZROZg06TVgl69OqCTZLPhiS91EoUCIysBkON0oq+2FgMmEzZdvgz45oHIbsdEbS2hn+1hdmhlqVTof6pdG7KycCwvjwC629+PLdeuLTpPQhiIdjrhrKqam2ClpfD46qBEJEJNQQE6e3rwc28vTkulaDUasb+nhwRke//v48cJ/f71oUaD731jd8k1AJrG/QMHkJmRgRy1Gnf5fGL7l0yGlORkbFGrMTw9DcuFC6BpGlEnT4KJjsZWlwu9Oh3Mo6M41daGS2wHeDyIrajAbGxsxE4IeyOqTkuDOj8fXxoMqBweRrLNhom6OhKQV1EBetUq+Ittk1qNAT4f3+3ciQ927MAXej2+ttnwWCaDRCxG7fXrULMzJQIToQBGRpAxO4vBixcxYbViTWMjGjZuxLF9+9DS0YHP+/rIbs6IxaiQStHe1YVcoxGe6moi5ysUmBaJsMluR//Zs0T2kkqF0QijPRTA+Dg4AgFma2qIcWxxMR6Xl0OUmIg0pRJ/CoVELhobw5OmJsy43ZC1tOCSXI77Dx5gc2vrHN0Mg2+2bUPe7t3oN5mQZTAAXG5IKkIBOJ2AQIARqRTrWApv3kRJbi7MY2NIPX+e5JssjwdDUinSUlNJatju+EyjwbdPUc1zOPBPdTUp6HcrK/FT8MFL3IQCmJgAUlJQI5Gg5PDhAGJZYyOarNagHfjTwApnZmYQV1YGzwKqC+PjcU6phGNqCkkaTaCr/I5CAbCHR3Q0XrNY8FtDQyBgXFERnPHxQQDWjY9jpKmJyH7s7sZ7nZ0hFHM9HgwcOoQN6emkXq52d/t2PqdKcTQahr0asdnhsWcAAAuA6MlJTOv1hD7T0BA2XLkS4pzNM3s4sXeIH8xm2NeuDdUBsNXhQK9eH/Y/ihvhzYj2ejF85AjWSyT4tLwct/y5D+tmcWHlmjV4OzubTFF2PkxOTpLLCYVTp+avZP4CoijEmkxwGgzwer3gq9Xw+Kp/8VDL04j4aqZISkK9XA5jZyfe9+Vtea6Xph0eAMPAqlJBGB+PzUeP4r5YvDRvK9AKC0BstWK0vp7Qzzt9OnAirsD/oiZhATRmZaEgLw8Xrl5FwcOHizp5FoV/AVkhFWhXQjINAAAAAElFTkSuQmCC"),
        // Please specify the base64 content of a 80x80 pixels image
        ExportMetadata("BigImageBase64", "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"),
        ExportMetadata("BackgroundColor", "Teal"),
        ExportMetadata("PrimaryFontColor", "White"),
        ExportMetadata("SecondaryFontColor", "Gold")]
    public class BulkDataFinderPlugin : PluginBase
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public BulkDataFinderPlugin()
        {
            // If you have external assemblies that you need to load, uncomment the following to
            // hook into the event that will fire when an Assembly fails to resolve
            // AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        public override IXrmToolBoxPluginControl GetControl()
        {
            return new BulkDataFinderControl();
        }

        /// <summary>
        /// Event fired by CLR when an assembly reference fails to load
        /// Assumes that related assemblies will be loaded from a subfolder named the same as the Plugin
        /// For example, a folder named Sample.XrmToolBox.MyPlugin
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            // base name of the assembly that failed to resolve
            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            // check to see if the failing assembly is one that we reference.
            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            // if the current unresolved assembly is referenced by our plugin, attempt to load
            if (refAssembly != null)
            {
                // load from the path to this plugin assembly, not host executable
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);

                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
                else
                {
                    throw new FileNotFoundException($"Unable to locate dependency: {assmbPath}");
                }
            }

            return loadAssembly;
        }
    }
}