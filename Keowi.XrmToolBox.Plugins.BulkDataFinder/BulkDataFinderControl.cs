using Keowi.XrmToolBox.Plugins.BulkDataFinder.AppCode;
using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Args;
using XrmToolBox.Extensibility.Interfaces;

namespace Keowi.XrmToolBox.Plugins.BulkDataFinder
{
    public partial class BulkDataFinderControl : PluginControlBase, IStatusBarMessenger, IGitHubPlugin
    {
        private const int WarningLimit = 100000;
        private SearchCriterias CurrentSearchCriterias;
        private List<string> CurrentValidAttributes;
        private Dictionary<string, int> EntitiesMetadata;
        private bool HasInputData;
        private bool HasSearchResults;
        private bool IsStopRequested;
        private MetadataManager metadataManager = null;
        private Settings mySettings;
        private List<Search> searchingDataList;
        private SearchResultsOptions SearchResultsOptions;

        public BulkDataFinderControl()
        {
            InitializeComponent();
        }

        public event EventHandler<StatusBarMessageEventArgs> SendMessageToStatusBar;

        public string RepositoryName => "Keowi.XrmToolBox.Plugins.BulkDataFinder";
        public string UserName => "WilliamKeo";

        /// <summary>
        /// This event occurs when the connection has been updated in XrmToolBox
        /// </summary>
        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (mySettings != null && detail != null)
            {
                mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
                LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
            }
        }

        private void allResultsRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            RenderResultsDetailsView(true);
        }

        private void attributesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            CurrentSearchCriterias.Attribute = (string)attributesComboBox.SelectedItem;

            searchButton.Enabled = true;
        }

        private void EnableControls(bool enable)
        {
            stopSearchToolStripButton.Enabled = !enable;
            openFileButton.Enabled = enable;
            ignoreHeaderCheckBox.Enabled = enable;
            entitiesComboBox.Enabled = enable;
            viewsComboBox.Enabled = enable;
            useFilteredViewCheckBox.Enabled = enable;
            attributesComboBox.Enabled = enable;
            recordIdRadioButton.Enabled = enable;
            recordIdAndPrimaryRadioButton.Enabled = enable;
            searchButton.Enabled = enable;
            tsbLoadMetadata.Enabled = enable;
            resultsDetailsGroupBox.Enabled = enable;
            exportResultsToolStripSplitButton.Enabled = enable;
            exportOnlyMatchingDataToolStripMenuItem.Enabled = enable;
        }

        private void entitiesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            CurrentSearchCriterias.Entity = (string)entitiesComboBox.SelectedItem;

            ExecuteMethod(GetSavedQureies);

            //TODO facto
            attributesComboBox.Items.Clear();
            attributesComboBox.ResetText();
            CurrentSearchCriterias.Attribute = null;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading attributes metadata...",
                AsyncArgument = CurrentSearchCriterias,
                Work = (worker, args) =>
                {
                    var criterias = (SearchCriterias)args.Argument;

                    args.Result = metadataManager.GetEntityPrimaryAndTextAttributes(criterias.Entity);
                },
                PostWorkCallBack = (args) =>
                {
                    var result = args.Result as Tuple<string, List<string>>;
                    var attrResult = result.Item2;
                    attributesComboBox.Items.AddRange(attrResult.ToArray());
                    CurrentValidAttributes = new List<string>(attrResult);

                    CurrentSearchCriterias.PrimaryAttribute = result.Item1;
                }
            });
        }

        private void exportOnlyMatchingDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExecuteMethod(ExportResults, false);
        }

        private void ExportResults(bool isFullResults)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Exporting search results...",
                AsyncArgument = SearchResultsOptions,
                Work = (worker, args) =>
                {
                    var criterias = (SearchResultsOptions)args.Argument;

                    ExcelPackage excelPkg = new ExcelPackage();
                    ExcelWorksheet worksheet = excelPkg.Workbook.Worksheets.Add("Search Results");

                    worksheet.Cells[1, 1].Value = $"{criterias.Attribute} ({criterias.Entity})";
                    worksheet.Cells[1, 1].Style.Font.Bold = true;
                    worksheet.Cells[1, 2].Value = "Processed";
                    worksheet.Cells[1, 2].Style.Font.Bold = true;
                    worksheet.Cells[1, 3].Value = "Found";
                    worksheet.Cells[1, 3].Style.Font.Bold = true;
                    worksheet.Cells[1, 4].Value = "Record Id";
                    worksheet.Cells[1, 4].Style.Font.Bold = true;
                    if (criterias.DisplayPrimaryAttribute)
                    {
                        worksheet.Cells[1, 5].Value = $"Primary attribute ({criterias.PrimaryAttribute})";
                        worksheet.Cells[1, 5].Style.Font.Bold = true;
                    }

                    var lineIndex = 2;

                    //Filtering output results.
                    var searchingDataOutput = searchingDataList;
                    if (!isFullResults)
                    {
                        searchingDataOutput = searchingDataOutput.Where(x => x.IsFound).ToList();
                    }

                    foreach (var searchItemResult in searchingDataOutput)
                    {
                        worksheet.Cells[lineIndex, 1].Value = searchItemResult.InputData;
                        worksheet.Cells[lineIndex, 2].Value = searchItemResult.IsProcessed;
                        worksheet.Cells[lineIndex, 3].Value = searchItemResult.IsFound;
                        if (searchItemResult.IsFound)
                        {
                            worksheet.Cells[lineIndex, 4].Value = searchItemResult.RecordId;
                        }
                        if (criterias.DisplayPrimaryAttribute)
                        {
                            worksheet.Cells[lineIndex, 5].Value = searchItemResult.PrimaryAttribute;
                        }

                        lineIndex++;
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    excelPkg.Save();

                    args.Result = excelPkg;
                },
                PostWorkCallBack = (args) =>
                {
                    var excelPkg = args.Result as ExcelPackage;

                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
                    //saveFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|Text files (*.txt)|*.txt|All files (*.*)|*.*";
                    saveFileDialog.FilterIndex = 0;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.FileName = $"BulkDataFinder Results {DateTime.Now.ToString("yyyyMMdd")}.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        using (FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.Create))
                        {
                            excelPkg.SaveAs(fs);
                        }
                    }

                    ShowInfoNotification($"The results has been saved to '{saveFileDialog.FileName}'", null, 20);
                }
            });
        }

        private void exportResultsToolStripButton_Click(object sender, EventArgs e)
        {
            ExecuteMethod(ExportResults, false);
        }

        private void exportResultsToolStripSplitButton_ButtonClick(object sender, EventArgs e)
        {
            ExecuteMethod(ExportResults, true);
        }

        private void GetEntitiesMetadata()
        {
            entitiesComboBox.Items.Clear();

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading entities metadata...",
                Work = (worker, args) =>
                {
                    var filteredEntities = metadataManager.GetEntitiesMetadata();

                    var entities = filteredEntities.ToDictionary(x => x.LogicalName, x => x.ObjectTypeCode.Value);
                    args.Result = entities;
                },
                PostWorkCallBack = (args) =>
                {
                    var result = args.Result as Dictionary<string, int>;
                    EntitiesMetadata = result;
                    entitiesComboBox.Items.AddRange(result.Keys.ToArray());
                }
            });
        }

        private void GetSavedQureies()
        {
            viewsComboBox.DataSource = null;
            //viewsComboBox.Items.Clear();
            CurrentSearchCriterias.FetchXml = null;

            var entityName = (string)entitiesComboBox.SelectedItem;
            var entityTypeCode = EntitiesMetadata.First(x => x.Key == entityName).Value;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading entity views...",
                AsyncArgument = entityTypeCode,
                Work = (worker, args) =>
                {
                    var entityCode = (int)args.Argument;

                    var savedQueries = metadataManager.GetSavedQueries(entityCode);

                    var userQueries = metadataManager.GetUserQueries(entityCode);

                    args.Result = new Tuple<EntityCollection, EntityCollection>(savedQueries, userQueries);
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var result = args.Result as Tuple<EntityCollection, EntityCollection>;
                    if (result != null)
                    {
                        var views = result.Item1.Entities.OrderBy(x => x.GetAttributeValue<string>("name")).Select(x => new ComboBoxItem
                        {
                            Key = x.GetAttributeValue<string>("name"),
                            Value = x.GetAttributeValue<string>("fetchxml")
                        }).ToList();
                        views.Insert(0, new ComboBoxItem { Key = "----- SYSTEM VIEWS -----", Value = "" });

                        if (result.Item2 != null && result.Item2.Entities.Any())
                        {
                            views.Add(new ComboBoxItem { Key = "----- PERSONAL VIEWS -----", Value = "" });
                            views.AddRange(result.Item2.Entities.OrderBy(x => x.GetAttributeValue<string>("name")).Select(x => new ComboBoxItem
                            {
                                Key = x.GetAttributeValue<string>("name"),
                                Value = x.GetAttributeValue<string>("fetchxml")
                            }).ToList());
                        }
                        viewsComboBox.DataSource = views.ToArray();
                        viewsComboBox.DisplayMember = "Key";
                        viewsComboBox.ValueMember = "Value";
                    }
                }
            });
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }

        private void ignoreHeaderCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            var curSearchingDataList = searchingDataList.Skip(ignoreHeaderCheckBox.Checked ? 1 : 0).ToList();
            rowNumberValue.Text = $"{curSearchingDataList.Count}";
        }

        private void matchingResultsRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            RenderResultsDetailsView(false);
        }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            //ShowInfoNotification("This is a notification that can lead to XrmToolBox repository", new Uri("https://github.com/MscrmTools/XrmToolBox"));

            // Loads or creates the settings for the plugin
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            {
                mySettings = new Settings();

                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }

            metadataManager = new MetadataManager(Service);

            searchButton.Enabled = false;

            CurrentSearchCriterias = new SearchCriterias
            {
                UseFilteredView = true
            };
            ExecuteMethod(GetEntitiesMetadata);

            ScintillaControl.InitXML(scintillaFetchXml, true);
        }

        /// <summary>
        /// This event occurs when the plugin is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyPluginControl_OnCloseTool(object sender, EventArgs e)
        {
            // Before leaving, save the settings
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        private void openFileButton_Click(object sender, EventArgs e)
        {
            HideNotification();

            //var fileContent = string.Empty;
            var filePath = string.Empty;

            searchingDataList = new List<Search>();

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
                //openFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    //var fileStream = openFileDialog.OpenFile();

                    //using (StreamReader reader = new StreamReader(fileStream))
                    //{
                    //    fileContent = reader.ReadToEnd();
                    //}

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var pck = new ExcelPackage(openFileDialog.OpenFile());
                    var worksheet = pck.Workbook.Worksheets[0];

                    for (var i = 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        var value = worksheet.Cells[i, 1].Value.ToString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            searchingDataList.Add(new Search
                            {
                                InputData = value
                            });
                        }
                    }
                }
            }

            var curSearchingDataList = searchingDataList.Skip(ignoreHeaderCheckBox.Checked ? 1 : 0).ToList();
            ShowInfoNotification($"{curSearchingDataList.Count} rows has been identified.", null, 20);
            rowNumberValue.Text = $"{curSearchingDataList.Count}";

            HasInputData = true;

            if (curSearchingDataList.Count > WarningLimit)
            {
                ShowWarningNotification("Your file contains more than 100000 rows, you might notice some delays during the search.", null, 20);
            }
        }

        private void RenderResultsDetailsView(bool isFullResults)
        {
            searchResultsListView.Items.Clear();

            var searchingDataOutput = searchingDataList.Skip(ignoreHeaderCheckBox.Checked ? 1 : 0).ToList();
            if (!isFullResults)
            {
                searchingDataOutput = searchingDataOutput.Where(x => x.IsFound).ToList();
            }

            foreach (var searchItemResult in searchingDataOutput)
            {
                var listItem = new ListViewItem(new[] {
                            searchItemResult.InputData,
                            searchItemResult.RecordId != Guid.Empty ? searchItemResult.RecordId.ToString() : string.Empty,
                            searchItemResult.PrimaryAttribute
                        });
                if (isFullResults && searchItemResult.IsFound)
                    listItem.ForeColor = Color.Green;
                if (isFullResults && !searchItemResult.IsProcessed)
                    listItem.BackColor = Color.LightGray;
                searchResultsListView.Items.Add(listItem);
            }
        }

        private void searchButton_Click(object sender, EventArgs e)
        {
            if (!HasInputData)
            {
                MessageBox.Show("No input file has been loaded for this search yet!",
                    "File not loaded", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!CurrentValidAttributes.Contains(CurrentSearchCriterias.Attribute))
            {
                MessageBox.Show("The attribute name specified is not valid for this search!",
                    "Invalid Search", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            CurrentSearchCriterias.FetchXml = (string)viewsComboBox.SelectedValue;
            if (CurrentSearchCriterias.UseFilteredView && string.IsNullOrEmpty(CurrentSearchCriterias.FetchXml))
            {
                MessageBox.Show("Please select a valid view before running the search!",
                    "Invalid Search", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            CurrentSearchCriterias.IgnoreHeader = ignoreHeaderCheckBox.Checked;
            CurrentSearchCriterias.DisplayPrimaryAttribute = recordIdAndPrimaryRadioButton.Checked;
            if (HasSearchResults
                && MessageBox.Show("Starting a new search will delete previous results. Do you want to continue?",
                    "Reset Results", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
            {
                return;
            }

            //Reset previous search results.
            HasSearchResults = false;
            rowNumberSearchedValue.Text = string.Empty;
            recordsFoundValue.Text = string.Empty;
            durationValue.Text = string.Empty;
            searchingDataList.ForEach(x =>
            {
                x.IsProcessed = false;
                x.IsFound = false;
                x.RecordId = Guid.Empty;
                x.PrimaryAttribute = string.Empty;
            });
            searchResultsListView.Items.Clear();

            IsStopRequested = false;
            EnableControls(false);

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Searching...",
                AsyncArgument = CurrentSearchCriterias,
                Work = (worker, args) =>
                {
                    var criterias = (SearchCriterias)args.Argument;

                    var sw = new Stopwatch();
                    sw.Start();

                    SendMessageToStatusBar?.Invoke(this, new StatusBarMessageEventArgs(0));

                    var curSearchingDataList = searchingDataList.Skip(criterias.IgnoreHeader ? 1 : 0).ToList();
                    Parallel.ForEach(curSearchingDataList,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    searchItem =>
                    {
                        if (IsStopRequested)
                            return;

                        var query = new QueryExpression(criterias.Entity);
                        // Use view filter
                        if (criterias.UseFilteredView)
                        {
                            var conversionRequest = new FetchXmlToQueryExpressionRequest
                            {
                                FetchXml = criterias.FetchXml
                            };
                            var conversionResponse =
                                (FetchXmlToQueryExpressionResponse)Service.Execute(conversionRequest);

                            query = conversionResponse.Query;
                            query.NoLock = true;

                            // Remove columns for performance.
                            query.ColumnSet = new ColumnSet();
                            foreach (var linkEntity in query.LinkEntities)
                            {
                                linkEntity.Columns = new ColumnSet();
                            }

                            query.Criteria.AddCondition(
                                new ConditionExpression(criterias.Attribute, ConditionOperator.Equal, searchItem.InputData));
                        }
                        else
                        {
                            query = new QueryExpression(criterias.Entity)
                            {
                                NoLock = true,
                                Criteria =
                                {
                                    Conditions =
                                    {
                                        new ConditionExpression(criterias.Attribute, ConditionOperator.Equal, searchItem.InputData)
                                    }
                                }
                            };
                        }

                        if (criterias.DisplayPrimaryAttribute)
                        {
                            query.ColumnSet = new ColumnSet(criterias.PrimaryAttribute);
                        }

                        var result = Service.RetrieveMultiple(query).Entities;
                        if (result.Any())
                        {
                            searchItem.IsFound = true;
                            searchItem.RecordId = result.First().Id;
                            if (criterias.DisplayPrimaryAttribute)
                            {
                                searchItem.PrimaryAttribute = result.First().GetAttributeValue<string>(criterias.PrimaryAttribute);
                            }
                        }
                        searchItem.IsProcessed = true;

                        //Managing progress bar.
                        var currentProgress = curSearchingDataList.Where(x => x.IsProcessed).Count() * 100 / curSearchingDataList.Count;
                        if (currentProgress % 10 == 0)
                        {
                            SendMessageToStatusBar?.Invoke(this, new StatusBarMessageEventArgs(currentProgress, string.Empty));
                        }
                    });

                    sw.Stop();

                    args.Result = sw.Elapsed;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    var duration = (TimeSpan)args.Result;

                    HasSearchResults = true;
                    EnableControls(true);

                    var curSearchingDataList = searchingDataList.Skip(ignoreHeaderCheckBox.Checked ? 1 : 0).ToList();
                    if (curSearchingDataList.All(x => x.IsProcessed))
                    {
                        MessageBox.Show("The search has completed successfully.\nPlease check the analysis results report and use the 'Export' action to download a copy.",
                            "Search completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("The search has been aborted! Unprocessed rows are hightlighted in gray.",
                            "Search interrupted", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    rowNumberSearchedValue.Text = $"{curSearchingDataList.Where(x => x.IsProcessed).Count()}";
                    recordsFoundValue.Text = $"{curSearchingDataList.Where(x => x.IsFound).Count()}";
                    durationValue.Text = $"{duration:dd\\.hh\\:mm\\:ss}";

                    SendMessageToStatusBar?.Invoke(this, new StatusBarMessageEventArgs(100, string.Empty));

                    RenderResultsDetailsView(true);

                    SearchResultsOptions = new SearchResultsOptions
                    {
                        Attribute = CurrentSearchCriterias.Attribute,
                        DisplayPrimaryAttribute = CurrentSearchCriterias.DisplayPrimaryAttribute,
                        Entity = CurrentSearchCriterias.Entity,
                        PrimaryAttribute = CurrentSearchCriterias.PrimaryAttribute
                    };
                }
            });
        }

        private void stopSearchToolStripButton_Click(object sender, EventArgs e)
        {
            IsStopRequested = true;
        }

        private void toolStripLabelDocumentationLink_Click(object sender, EventArgs e)
        {
            Process.Start("https://williamkeo293423625.wordpress.com/2021/08/25/xrmtoolbox-new-plugin-bulk-data-finder/");
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }

        private void tsbLoadMetadata_Click(object sender, EventArgs e)
        {
            ExecuteMethod(GetEntitiesMetadata);
        }

        private void useFilteredViewCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (!useFilteredViewCheckBox.Checked)
            {
                var result = MessageBox.Show("The search will be executed on the whole entity.", "Search Perimeter",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Cancel)
                {
                    useFilteredViewCheckBox.Checked = true;
                }
            }
            CurrentSearchCriterias.UseFilteredView = useFilteredViewCheckBox.Checked;
            viewsComboBox.Enabled = useFilteredViewCheckBox.Checked;
        }

        private void viewsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (viewsComboBox.SelectedItem != null)
            {
                scintillaFetchXml.Text = ((ComboBoxItem)viewsComboBox.SelectedItem).Value;
            }
            else
            {
                scintillaFetchXml.Text = "";
            }
            ScintillaControl.Format(scintillaFetchXml);
        }
    }
}