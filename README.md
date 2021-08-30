# Keowi.XrmToolBox.Plugins.BulkDataFinder
This plugin allows you to perform bulk search on text fields within Microsoft Dynamics 365 or Dataverse from your external data.

For more details, please check this [article](https://williamkeo293423625.wordpress.com/2021/08/25/xrmtoolbox-new-plugin-bulk-data-finder/).

## How to use
#### 1. Select your input file
Single column Excel file is required at the moment. Each row will contains the text input you are looking for. You can ignore the file header using the checkbox.
#### 2. Define search settings
Choose the entity, a filtered view and the attribute on which the search will be looking at.
The view selection enables data filtering based on the systems views available on the entity but also personal user views defined on the logged account.
If you are attempting a search on the whole entity uncheck the "Use Filtered Views" options. Keep in mind that performance can be affected by searching into a large table.
#### 3. Choose your results options
Decide whether you want the primary attribute value of the entity or not. It might be useful for those need more information than the record Id.
#### 4. Run the search
Once the previous steps configured, you can now run the search by hitting the button. You can interupt the search at any time when necessary using the "Stop Search" command.
#### 5. Export the results
When the search is completed, you can look into the results on the screen but also exported the detailled analysis Excel file for further purposes. Multiple export options are proposed according to your needs.
