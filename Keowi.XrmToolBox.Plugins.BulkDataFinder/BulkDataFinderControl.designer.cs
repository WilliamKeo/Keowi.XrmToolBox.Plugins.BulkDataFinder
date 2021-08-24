
namespace Keowi.XrmToolBox.Plugins.BulkDataFinder
{
    partial class BulkDataFinderControl
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BulkDataFinderControl));
            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
            this.tsbClose = new System.Windows.Forms.ToolStripButton();
            this.tssSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbLoadMetadata = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.stopSearchToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.exportResultsToolStripSplitButton = new System.Windows.Forms.ToolStripSplitButton();
            this.exportOnlyMatchingDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ignoreHeaderCheckBox = new System.Windows.Forms.CheckBox();
            this.openFileButton = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.useFilteredViewCheckBox = new System.Windows.Forms.CheckBox();
            this.viewsComboBox = new System.Windows.Forms.ComboBox();
            this.searchButton = new System.Windows.Forms.Button();
            this.attributesComboBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.entitiesComboBox = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.rowNumberSearchedValue = new System.Windows.Forms.Label();
            this.rowNumberSearchedLabel = new System.Windows.Forms.Label();
            this.durationValue = new System.Windows.Forms.Label();
            this.recordsFoundValue = new System.Windows.Forms.Label();
            this.rowNumberValue = new System.Windows.Forms.Label();
            this.durationLabel = new System.Windows.Forms.Label();
            this.recordsFoundLabel = new System.Windows.Forms.Label();
            this.rowNumberLabel = new System.Windows.Forms.Label();
            this.searchResultsListView = new System.Windows.Forms.ListView();
            this.searchInput = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.recordId = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.resultsDetailsGroupBox = new System.Windows.Forms.GroupBox();
            this.matchingResultsRadioButton = new System.Windows.Forms.RadioButton();
            this.allResultsRadioButton = new System.Windows.Forms.RadioButton();
            this.toolStripLabelDocumentationLink = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenu.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.resultsDetailsGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripMenu
            // 
            this.toolStripMenu.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbClose,
            this.tssSeparator1,
            this.tsbLoadMetadata,
            this.toolStripSeparator1,
            this.stopSearchToolStripButton,
            this.toolStripSeparator2,
            this.exportResultsToolStripSplitButton,
            this.toolStripSeparator3,
            this.toolStripLabelDocumentationLink});
            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripMenu.Name = "toolStripMenu";
            this.toolStripMenu.Size = new System.Drawing.Size(1216, 25);
            this.toolStripMenu.TabIndex = 4;
            this.toolStripMenu.Text = "toolStrip1";
            // 
            // tsbClose
            // 
            this.tsbClose.Image = ((System.Drawing.Image)(resources.GetObject("tsbClose.Image")));
            this.tsbClose.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.tsbClose.Name = "tsbClose";
            this.tsbClose.Size = new System.Drawing.Size(56, 22);
            this.tsbClose.Text = "Close";
            this.tsbClose.Click += new System.EventHandler(this.tsbClose_Click);
            // 
            // tssSeparator1
            // 
            this.tssSeparator1.Name = "tssSeparator1";
            this.tssSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // tsbLoadMetadata
            // 
            this.tsbLoadMetadata.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbLoadMetadata.Name = "tsbLoadMetadata";
            this.tsbLoadMetadata.Size = new System.Drawing.Size(90, 22);
            this.tsbLoadMetadata.Text = "Load metadata";
            this.tsbLoadMetadata.ToolTipText = "Load metadata";
            this.tsbLoadMetadata.Click += new System.EventHandler(this.tsbLoadMetadata_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // stopSearchToolStripButton
            // 
            this.stopSearchToolStripButton.Enabled = false;
            this.stopSearchToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("stopSearchToolStripButton.Image")));
            this.stopSearchToolStripButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.stopSearchToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.stopSearchToolStripButton.Name = "stopSearchToolStripButton";
            this.stopSearchToolStripButton.Size = new System.Drawing.Size(89, 22);
            this.stopSearchToolStripButton.Text = "Stop Search";
            this.stopSearchToolStripButton.ToolTipText = "Stop Search";
            this.stopSearchToolStripButton.Click += new System.EventHandler(this.stopSearchToolStripButton_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // exportResultsToolStripSplitButton
            // 
            this.exportResultsToolStripSplitButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exportOnlyMatchingDataToolStripMenuItem});
            this.exportResultsToolStripSplitButton.Enabled = false;
            this.exportResultsToolStripSplitButton.Image = ((System.Drawing.Image)(resources.GetObject("exportResultsToolStripSplitButton.Image")));
            this.exportResultsToolStripSplitButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.exportResultsToolStripSplitButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.exportResultsToolStripSplitButton.Name = "exportResultsToolStripSplitButton";
            this.exportResultsToolStripSplitButton.Size = new System.Drawing.Size(135, 22);
            this.exportResultsToolStripSplitButton.Text = "Export Full Results";
            this.exportResultsToolStripSplitButton.ToolTipText = "Export Full Results";
            this.exportResultsToolStripSplitButton.ButtonClick += new System.EventHandler(this.exportResultsToolStripSplitButton_ButtonClick);
            // 
            // exportOnlyMatchingDataToolStripMenuItem
            // 
            this.exportOnlyMatchingDataToolStripMenuItem.Enabled = false;
            this.exportOnlyMatchingDataToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("exportOnlyMatchingDataToolStripMenuItem.Image")));
            this.exportOnlyMatchingDataToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.exportOnlyMatchingDataToolStripMenuItem.Name = "exportOnlyMatchingDataToolStripMenuItem";
            this.exportOnlyMatchingDataToolStripMenuItem.Size = new System.Drawing.Size(217, 24);
            this.exportOnlyMatchingDataToolStripMenuItem.Text = "Export Only Matching Data";
            this.exportOnlyMatchingDataToolStripMenuItem.Click += new System.EventHandler(this.exportOnlyMatchingDataToolStripMenuItem_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ignoreHeaderCheckBox);
            this.groupBox1.Controls.Add(this.openFileButton);
            this.groupBox1.Location = new System.Drawing.Point(3, 28);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(656, 100);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Input Data";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // ignoreHeaderCheckBox
            // 
            this.ignoreHeaderCheckBox.AutoSize = true;
            this.ignoreHeaderCheckBox.Location = new System.Drawing.Point(20, 56);
            this.ignoreHeaderCheckBox.Name = "ignoreHeaderCheckBox";
            this.ignoreHeaderCheckBox.Size = new System.Drawing.Size(108, 17);
            this.ignoreHeaderCheckBox.TabIndex = 1;
            this.ignoreHeaderCheckBox.Text = "Ignore file header";
            this.ignoreHeaderCheckBox.UseVisualStyleBackColor = true;
            this.ignoreHeaderCheckBox.CheckedChanged += new System.EventHandler(this.ignoreHeaderCheckBox_CheckedChanged);
            // 
            // openFileButton
            // 
            this.openFileButton.Image = ((System.Drawing.Image)(resources.GetObject("openFileButton.Image")));
            this.openFileButton.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.openFileButton.Location = new System.Drawing.Point(20, 20);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(120, 25);
            this.openFileButton.TabIndex = 0;
            this.openFileButton.Text = "Open File";
            this.openFileButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.openFileButton_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.useFilteredViewCheckBox);
            this.groupBox2.Controls.Add(this.viewsComboBox);
            this.groupBox2.Controls.Add(this.searchButton);
            this.groupBox2.Controls.Add(this.attributesComboBox);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.entitiesComboBox);
            this.groupBox2.Location = new System.Drawing.Point(3, 134);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(656, 126);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Search Settings";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "View";
            // 
            // useFilteredViewCheckBox
            // 
            this.useFilteredViewCheckBox.AutoSize = true;
            this.useFilteredViewCheckBox.Checked = true;
            this.useFilteredViewCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.useFilteredViewCheckBox.Location = new System.Drawing.Point(312, 49);
            this.useFilteredViewCheckBox.Name = "useFilteredViewCheckBox";
            this.useFilteredViewCheckBox.Size = new System.Drawing.Size(108, 17);
            this.useFilteredViewCheckBox.TabIndex = 6;
            this.useFilteredViewCheckBox.Text = "Use Filtered View";
            this.useFilteredViewCheckBox.UseVisualStyleBackColor = true;
            this.useFilteredViewCheckBox.CheckedChanged += new System.EventHandler(this.useFilteredViewCheckBox_CheckedChanged);
            // 
            // viewsComboBox
            // 
            this.viewsComboBox.FormattingEnabled = true;
            this.viewsComboBox.Location = new System.Drawing.Point(105, 46);
            this.viewsComboBox.Name = "viewsComboBox";
            this.viewsComboBox.Size = new System.Drawing.Size(200, 21);
            this.viewsComboBox.TabIndex = 5;
            this.viewsComboBox.SelectedIndexChanged += new System.EventHandler(this.viewsComboBox_SelectedIndexChanged);
            // 
            // searchButton
            // 
            this.searchButton.Enabled = false;
            this.searchButton.Location = new System.Drawing.Point(500, 97);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new System.Drawing.Size(150, 23);
            this.searchButton.TabIndex = 4;
            this.searchButton.Text = "Start Search";
            this.searchButton.UseVisualStyleBackColor = true;
            this.searchButton.Click += new System.EventHandler(this.searchButton_Click);
            // 
            // attributesComboBox
            // 
            this.attributesComboBox.FormattingEnabled = true;
            this.attributesComboBox.Location = new System.Drawing.Point(105, 73);
            this.attributesComboBox.Name = "attributesComboBox";
            this.attributesComboBox.Size = new System.Drawing.Size(200, 21);
            this.attributesComboBox.TabIndex = 3;
            this.attributesComboBox.SelectedIndexChanged += new System.EventHandler(this.attributesComboBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Attribute";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Entity";
            // 
            // entitiesComboBox
            // 
            this.entitiesComboBox.FormattingEnabled = true;
            this.entitiesComboBox.Location = new System.Drawing.Point(105, 19);
            this.entitiesComboBox.Name = "entitiesComboBox";
            this.entitiesComboBox.Size = new System.Drawing.Size(200, 21);
            this.entitiesComboBox.TabIndex = 0;
            this.entitiesComboBox.SelectedIndexChanged += new System.EventHandler(this.entitiesComboBox_SelectedIndexChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.rowNumberSearchedValue);
            this.groupBox3.Controls.Add(this.rowNumberSearchedLabel);
            this.groupBox3.Controls.Add(this.durationValue);
            this.groupBox3.Controls.Add(this.recordsFoundValue);
            this.groupBox3.Controls.Add(this.rowNumberValue);
            this.groupBox3.Controls.Add(this.durationLabel);
            this.groupBox3.Controls.Add(this.recordsFoundLabel);
            this.groupBox3.Controls.Add(this.rowNumberLabel);
            this.groupBox3.Location = new System.Drawing.Point(3, 266);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(656, 129);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Analysis Results";
            // 
            // rowNumberSearchedValue
            // 
            this.rowNumberSearchedValue.AutoSize = true;
            this.rowNumberSearchedValue.Location = new System.Drawing.Point(156, 47);
            this.rowNumberSearchedValue.MinimumSize = new System.Drawing.Size(100, 0);
            this.rowNumberSearchedValue.Name = "rowNumberSearchedValue";
            this.rowNumberSearchedValue.Size = new System.Drawing.Size(100, 13);
            this.rowNumberSearchedValue.TabIndex = 7;
            // 
            // rowNumberSearchedLabel
            // 
            this.rowNumberSearchedLabel.AutoSize = true;
            this.rowNumberSearchedLabel.Location = new System.Drawing.Point(17, 47);
            this.rowNumberSearchedLabel.Name = "rowNumberSearchedLabel";
            this.rowNumberSearchedLabel.Size = new System.Drawing.Size(131, 13);
            this.rowNumberSearchedLabel.TabIndex = 6;
            this.rowNumberSearchedLabel.Text = "Number of rows searched:";
            // 
            // durationValue
            // 
            this.durationValue.AutoSize = true;
            this.durationValue.Location = new System.Drawing.Point(156, 93);
            this.durationValue.MinimumSize = new System.Drawing.Size(100, 0);
            this.durationValue.Name = "durationValue";
            this.durationValue.Size = new System.Drawing.Size(100, 13);
            this.durationValue.TabIndex = 5;
            // 
            // recordsFoundValue
            // 
            this.recordsFoundValue.AutoSize = true;
            this.recordsFoundValue.Location = new System.Drawing.Point(156, 70);
            this.recordsFoundValue.MinimumSize = new System.Drawing.Size(100, 0);
            this.recordsFoundValue.Name = "recordsFoundValue";
            this.recordsFoundValue.Size = new System.Drawing.Size(100, 13);
            this.recordsFoundValue.TabIndex = 4;
            // 
            // rowNumberValue
            // 
            this.rowNumberValue.AutoSize = true;
            this.rowNumberValue.Location = new System.Drawing.Point(156, 25);
            this.rowNumberValue.MinimumSize = new System.Drawing.Size(100, 0);
            this.rowNumberValue.Name = "rowNumberValue";
            this.rowNumberValue.Size = new System.Drawing.Size(100, 13);
            this.rowNumberValue.TabIndex = 3;
            // 
            // durationLabel
            // 
            this.durationLabel.AutoSize = true;
            this.durationLabel.Location = new System.Drawing.Point(17, 93);
            this.durationLabel.Name = "durationLabel";
            this.durationLabel.Size = new System.Drawing.Size(50, 13);
            this.durationLabel.TabIndex = 2;
            this.durationLabel.Text = "Duration:";
            // 
            // recordsFoundLabel
            // 
            this.recordsFoundLabel.AutoSize = true;
            this.recordsFoundLabel.Location = new System.Drawing.Point(17, 70);
            this.recordsFoundLabel.Name = "recordsFoundLabel";
            this.recordsFoundLabel.Size = new System.Drawing.Size(80, 13);
            this.recordsFoundLabel.TabIndex = 1;
            this.recordsFoundLabel.Text = "Records found:";
            // 
            // rowNumberLabel
            // 
            this.rowNumberLabel.AutoSize = true;
            this.rowNumberLabel.Location = new System.Drawing.Point(17, 25);
            this.rowNumberLabel.Name = "rowNumberLabel";
            this.rowNumberLabel.Size = new System.Drawing.Size(84, 13);
            this.rowNumberLabel.TabIndex = 0;
            this.rowNumberLabel.Text = "Number of rows:";
            // 
            // searchResultsListView
            // 
            this.searchResultsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.searchInput,
            this.recordId});
            this.searchResultsListView.GridLines = true;
            this.searchResultsListView.Location = new System.Drawing.Point(6, 56);
            this.searchResultsListView.MultiSelect = false;
            this.searchResultsListView.Name = "searchResultsListView";
            this.searchResultsListView.Size = new System.Drawing.Size(512, 305);
            this.searchResultsListView.TabIndex = 9;
            this.searchResultsListView.UseCompatibleStateImageBehavior = false;
            this.searchResultsListView.View = System.Windows.Forms.View.Details;
            // 
            // searchInput
            // 
            this.searchInput.Text = "Search Input";
            this.searchInput.Width = 180;
            // 
            // recordId
            // 
            this.recordId.Text = "Record Id";
            this.recordId.Width = 220;
            // 
            // resultsDetailsGroupBox
            // 
            this.resultsDetailsGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.resultsDetailsGroupBox.Controls.Add(this.matchingResultsRadioButton);
            this.resultsDetailsGroupBox.Controls.Add(this.allResultsRadioButton);
            this.resultsDetailsGroupBox.Controls.Add(this.searchResultsListView);
            this.resultsDetailsGroupBox.Enabled = false;
            this.resultsDetailsGroupBox.Location = new System.Drawing.Point(665, 28);
            this.resultsDetailsGroupBox.Name = "resultsDetailsGroupBox";
            this.resultsDetailsGroupBox.Size = new System.Drawing.Size(548, 367);
            this.resultsDetailsGroupBox.TabIndex = 10;
            this.resultsDetailsGroupBox.TabStop = false;
            this.resultsDetailsGroupBox.Text = "Results Details";
            // 
            // matchingResultsRadioButton
            // 
            this.matchingResultsRadioButton.AutoSize = true;
            this.matchingResultsRadioButton.Location = new System.Drawing.Point(98, 19);
            this.matchingResultsRadioButton.Name = "matchingResultsRadioButton";
            this.matchingResultsRadioButton.Size = new System.Drawing.Size(131, 17);
            this.matchingResultsRadioButton.TabIndex = 11;
            this.matchingResultsRadioButton.Text = "Only Matching Results";
            this.matchingResultsRadioButton.UseVisualStyleBackColor = true;
            this.matchingResultsRadioButton.CheckedChanged += new System.EventHandler(this.matchingResultsRadioButton_CheckedChanged);
            // 
            // allResultsRadioButton
            // 
            this.allResultsRadioButton.AutoSize = true;
            this.allResultsRadioButton.Checked = true;
            this.allResultsRadioButton.Location = new System.Drawing.Point(7, 20);
            this.allResultsRadioButton.Name = "allResultsRadioButton";
            this.allResultsRadioButton.Size = new System.Drawing.Size(74, 17);
            this.allResultsRadioButton.TabIndex = 10;
            this.allResultsRadioButton.TabStop = true;
            this.allResultsRadioButton.Text = "All Results";
            this.allResultsRadioButton.UseVisualStyleBackColor = true;
            this.allResultsRadioButton.CheckedChanged += new System.EventHandler(this.allResultsRadioButton_CheckedChanged);
            // 
            // toolStripLabelDocumentationLink
            // 
            this.toolStripLabelDocumentationLink.Image = ((System.Drawing.Image)(resources.GetObject("toolStripLabelDocumentationLink.Image")));
            this.toolStripLabelDocumentationLink.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.toolStripLabelDocumentationLink.IsLink = true;
            this.toolStripLabelDocumentationLink.LinkColor = System.Drawing.Color.Teal;
            this.toolStripLabelDocumentationLink.Name = "toolStripLabelDocumentationLink";
            this.toolStripLabelDocumentationLink.Size = new System.Drawing.Size(106, 22);
            this.toolStripLabelDocumentationLink.Text = "Documentation";
            this.toolStripLabelDocumentationLink.Click += new System.EventHandler(this.toolStripLabelDocumentationLink_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
            // 
            // BulkDataFinderControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.resultsDetailsGroupBox);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.toolStripMenu);
            this.Name = "BulkDataFinderControl";
            this.PluginIcon = ((System.Drawing.Icon)(resources.GetObject("$this.PluginIcon")));
            this.Size = new System.Drawing.Size(1216, 540);
            this.TabIcon = ((System.Drawing.Image)(resources.GetObject("$this.TabIcon")));
            this.Load += new System.EventHandler(this.MyPluginControl_Load);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.resultsDetailsGroupBox.ResumeLayout(false);
            this.resultsDetailsGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolStrip toolStripMenu;
        private System.Windows.Forms.ToolStripButton tsbClose;
        private System.Windows.Forms.ToolStripSeparator tssSeparator1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ComboBox attributesComboBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox entitiesComboBox;
        private System.Windows.Forms.Button searchButton;
        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.Label recordsFoundLabel;
        private System.Windows.Forms.Label rowNumberLabel;
        private System.Windows.Forms.ToolStripButton tsbLoadMetadata;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.Label durationLabel;
        private System.Windows.Forms.Label durationValue;
        private System.Windows.Forms.Label recordsFoundValue;
        private System.Windows.Forms.Label rowNumberValue;
        private System.Windows.Forms.ToolStripButton stopSearchToolStripButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.Label rowNumberSearchedValue;
        private System.Windows.Forms.Label rowNumberSearchedLabel;
        private System.Windows.Forms.ToolStripSplitButton exportResultsToolStripSplitButton;
        private System.Windows.Forms.ToolStripMenuItem exportOnlyMatchingDataToolStripMenuItem;
        private System.Windows.Forms.ListView searchResultsListView;
        private System.Windows.Forms.ColumnHeader searchInput;
        private System.Windows.Forms.ColumnHeader recordId;
        private System.Windows.Forms.GroupBox resultsDetailsGroupBox;
        private System.Windows.Forms.RadioButton matchingResultsRadioButton;
        private System.Windows.Forms.RadioButton allResultsRadioButton;
        private System.Windows.Forms.ComboBox viewsComboBox;
        private System.Windows.Forms.CheckBox useFilteredViewCheckBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox ignoreHeaderCheckBox;
        private System.Windows.Forms.ToolStripLabel toolStripLabelDocumentationLink;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
    }
}
