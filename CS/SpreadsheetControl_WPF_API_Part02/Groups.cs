using SpreadsheetControl_API_Part02;
using System.Collections.Generic;

namespace SpreadsheetControl_WPF_API_Part02
{
    public partial class Groups : List<Group>
    {
        public static Groups InitData()
        {
            Groups examples = new Groups();

            #region GroupNodes
            examples.Add(new Group("Shapes"));
            examples.Add(new Group("Custom Functions"));
            examples.Add(new Group("Tables"));
            examples.Add(new Group("Protection"));
            examples.Add(new Group("Search"));
            examples.Add(new Group("Sort"));
            examples.Add(new Group("Export"));
            examples.Add(new Group("Group Data"));
            examples.Add(new Group("Filter Data"));
            examples.Add(new Group("Document Properties"));
            examples.Add(new Group("Custom Xml Parts"));
            examples.Add(new Group("Data Validation"));
            examples.Add(new Group("Rows and Columns"));
            examples.Add(new Group("Form Controls"));

            #endregion

            #region ExampleNodes
            // Add nodes to the "Shapes" group of examples.
            examples[0].Items.Add(new SpreadsheetExample("Insert a picture", ShapeActions.InsertShapeAction));
            examples[0].Items.Add(new SpreadsheetExample("Insert a picture from URI", ShapeActions.InsertShapeFromUriAction));
            examples[0].Items.Add(new SpreadsheetExample("Modify a picture", ShapeActions.ModifyShapeAction));

            // Add nodes to the "Custom Functions" group of examples.
            examples[1].Items.Add(new SpreadsheetExample("Add a SPHEREMASS function", CustomFunctionActions.SphereMassAction));

            // Add nodes to the "Tables" group of examples.
            examples[2].Items.Add(new SpreadsheetExample("Create a table", TableActions.CreateTableAction));
            examples[2].Items.Add(new SpreadsheetExample("Apply a custom style", TableActions.CustomTableStyleAction));

            // Add nodes to the "Protection" group of examples.
            examples[3].Items.Add(new SpreadsheetExample("Protect workbook", ProtectionActions.ProtectWorkbookAction));
            examples[3].Items.Add(new SpreadsheetExample("Protect worksheet", ProtectionActions.ProtectWorksheetAction));
            examples[3].Items.Add(new SpreadsheetExample("Protect range", ProtectionActions.ProtectRangeAction));


            // Add nodes to the "Search" group of examples.
            examples[4].Items.Add(new SpreadsheetExample("Simple Search", SearchActions.SimpleSearchAction));
            examples[4].Items.Add(new SpreadsheetExample("Advanced Search", SearchActions.AdvancedSearchAction));

            // Add nodes to the "Sort" group of examples.
            examples[5].Items.Add(new SpreadsheetExample("Simple sort", SortActions.SimpleSortAction));
            examples[5].Items.Add(new SpreadsheetExample("Sort in descending order", SortActions.DescendingOrderAction));
            examples[5].Items.Add(new SpreadsheetExample("Sort using custom comparer", SortActions.SelectComparerAction));
            examples[5].Items.Add(new SpreadsheetExample("Sort by column", SortActions.SortBySpecifiedColumnAction));
            examples[5].Items.Add(new SpreadsheetExample("Sort by multiple columns", SortActions.SortByMultipleColumnsAction));
            examples[5].Items.Add(new SpreadsheetExample("Sort by fill", SortActions.SortByFillColorAction));
            examples[5].Items.Add(new SpreadsheetExample("Sort by font color", SortActions.SortByFontColorAction));

            // Add nodes to the "Export" group of examples.
            examples[6].Items.Add(new SpreadsheetExample("Export to HTML", ExportActions.ExportToHTMLAction));

            // Add nodes to the "Group Data" group of examples.
            examples[7].Items.Add(new SpreadsheetExample("Group Rows", GroupAndOutlineActions.GroupRowsAction));
            examples[7].Items.Add(new SpreadsheetExample("Group Columns", GroupAndOutlineActions.GroupColumnsAction));
            examples[7].Items.Add(new SpreadsheetExample("Unroup Rows", GroupAndOutlineActions.UngroupRowsAction));
            examples[7].Items.Add(new SpreadsheetExample("Unroup Columns", GroupAndOutlineActions.UngroupColumnsAction));
            examples[7].Items.Add(new SpreadsheetExample("Create an Auto Outline", GroupAndOutlineActions.AutoOutlineAction));
            examples[7].Items.Add(new SpreadsheetExample("Insert Subtotals", GroupAndOutlineActions.SubtotalAction));

            // Add nodes to the "Filter Data" group of examples.
            examples[8].Items.Add(new SpreadsheetExample("Enable filtering", AutoFilterActions.ApplyFilterAction));
            examples[8].Items.Add(new SpreadsheetExample("Sort by single column", AutoFilterActions.FilterAndSortBySingleColumnAction));
            examples[8].Items.Add(new SpreadsheetExample("Sort by multiple columns", AutoFilterActions.FilterAndSortByMultipleColumnsAction));
            examples[8].Items.Add(new SpreadsheetExample("Custom number filter", AutoFilterActions.FilterNumericByConditionAction));
            examples[8].Items.Add(new SpreadsheetExample("Custom text filter", AutoFilterActions.FilterTextByConditionAction));
            examples[8].Items.Add(new SpreadsheetExample("Custom date filter", AutoFilterActions.FilterDatesByConditionAction));
            examples[8].Items.Add(new SpreadsheetExample("Filter by single value", AutoFilterActions.FilterByValuesAction));
            examples[8].Items.Add(new SpreadsheetExample("Filter by multiple values", AutoFilterActions.FilterByMultipleValuesAction));
            examples[8].Items.Add(new SpreadsheetExample("Filter mixed data types by values", AutoFilterActions.FilterMixedDataTypesByValuesAction));
            examples[8].Items.Add(new SpreadsheetExample("Apply Top 10 filter", AutoFilterActions.Top10FilterAction));
            examples[8].Items.Add(new SpreadsheetExample("Apply dynamic filter", AutoFilterActions.DynamicFilterAction));
            examples[8].Items.Add(new SpreadsheetExample("Sort and filter by color", AutoFilterActions.FilterAndSortByColorAction));
            examples[8].Items.Add(new SpreadsheetExample("Filter by font color", AutoFilterActions.FilterByFontColorAction));
            examples[8].Items.Add(new SpreadsheetExample("Filter by fill color", AutoFilterActions.FilterByFillColorAction));
            examples[8].Items.Add(new SpreadsheetExample("Filter by background color", AutoFilterActions.FilterByBackgroundColorAction));
            examples[8].Items.Add(new SpreadsheetExample("Reapply filter", AutoFilterActions.ReapplyFilterAction));
            examples[8].Items.Add(new SpreadsheetExample("Clear filter", AutoFilterActions.ClearFilterAction));
            examples[8].Items.Add(new SpreadsheetExample("Disable filtering", AutoFilterActions.DisableFilterAction));

            // Add nodes to the "Document Properties" group of examples.
            examples[9].Items.Add(new SpreadsheetExample("Built-in properties", DocumentPropertiesActions.BuiltInPropertiesAction));
            examples[9].Items.Add(new SpreadsheetExample("Custom properties", DocumentPropertiesActions.CustomPropertiesAction));

            // Add nodes to the "Custom Xml" group of examples.
            examples[10].Items.Add(new SpreadsheetExample("Store Custom XML Parts", CustomXmlPartActions.StoreCustomXmlPartAction));
            examples[10].Items.Add(new SpreadsheetExample("Obtain Custom XML Parts", CustomXmlPartActions.ObtainCustomXmlPartAction));
            examples[10].Items.Add(new SpreadsheetExample("Modify Custom XML Parts", CustomXmlPartActions.ModifyCustomXmlPartAction));

            // Add nodes to the "Data Validation" group of examples.
            examples[11].Items.Add(new SpreadsheetExample("Add Data Validation", DataValidationActions.AddDataValidationAction));
            examples[11].Items.Add(new SpreadsheetExample("Change Validation Criteria", DataValidationActions.ChangeCriteriaAction));
            examples[11].Items.Add(new SpreadsheetExample("Get Data Validation", DataValidationActions.GetDataValidationAction));
            examples[11].Items.Add(new SpreadsheetExample("Remove All Data Validations", DataValidationActions.RemoveAllDataValidationsAction));
            examples[11].Items.Add(new SpreadsheetExample("Remove Data Validation", DataValidationActions.RemoveDataValidationAction));
            examples[11].Items.Add(new SpreadsheetExample("Show Error Message", DataValidationActions.ShowErrorMessageAction));
            examples[11].Items.Add(new SpreadsheetExample("Show Input Message", DataValidationActions.ShowInputMessageAction));
            examples[11].Items.Add(new SpreadsheetExample("Use Union Range", DataValidationActions.UseUnionRangeAction));

            // Add nodes to the "Rows and Columns" group of examples.
            examples[12].Items.Add(new SpreadsheetExample("Delete Columns Based On Condition", RowAndColumnActions.DeleteColumnsBasedOnConditionAction));
            examples[12].Items.Add(new SpreadsheetExample("Delete Rows Based On Condition", RowAndColumnActions.DeleteRowsBasedOnConditionAction));

            // Add nodes to the "Form Controls" group of examples.
            examples[13].Items.Add(new SpreadsheetExample("Create Form Controls", FormControlActions.CreateFormControlsAction));
            examples[13].Items.Add(new SpreadsheetExample("Edit Form Controls", FormControlActions.EditFormControlsAction));
            return examples;
            #endregion
        }
    }
}
