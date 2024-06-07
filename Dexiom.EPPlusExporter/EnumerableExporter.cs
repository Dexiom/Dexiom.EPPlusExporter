using System;
using System.Collections.Generic;
using System.Linq;
using Dexiom.EPPlusExporter.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter
{
    #region Create Method (using type inference)
    public static class EnumerableExporter
    {
        public static EnumerableExporter<T> Create<T>(IEnumerable<T> data, TableStyles tableStyles = TableStyles.Light1) where T : class => new EnumerableExporter<T>(data) { TableStyle = tableStyles };
        public static EnumerableExporter<T> Create<T>(IEnumerable<T> data, IEnumerable<DynamicProperty<T>> dynamicProperties, TableStyles tableStyles = TableStyles.Light1) where T : class => new EnumerableExporter<T>(data, dynamicProperties) { TableStyle = tableStyles };
    }
    #endregion

    public class EnumerableExporter<T> : TableExporter<T>
        where T : class
    {
        public IEnumerable<T> Data { get; set; }
        public IEnumerable<DynamicProperty<T>> DynamicProperties { get; set; }

        #region Constructors
        public EnumerableExporter(IEnumerable<T> data)
        {
            Data = data;
        }

        public EnumerableExporter(IEnumerable<T> data, IEnumerable<DynamicProperty<T>> dynamicProperties)
        {
            Data = data;
            DynamicProperties = dynamicProperties;
        }
        #endregion

        #region Protected
        protected override ExcelRange AddWorksheet(ExcelPackage package)
        {
            const int headerFirstRow = 1;
            const int headerFirstCol = 1;
            const int dataFirstRow = 2;
            const int dataFirstCol = 1;

            if (Data == null)
                return null;

            //let's avoid multiple enumeration
            var data = Data as IList<T> ?? Data.ToList();
            IList<DynamicProperty<T>> dynamicProperties = null;
            if (DynamicProperties != null)
                dynamicProperties = DynamicProperties as IList<DynamicProperty<T>> ?? DynamicProperties.ToList();

            //get available properties
            var properties = ReflectionHelper.GetBaseTypeOfEnumerable(data).GetProperties()
                .Where(p => !IgnoredProperties.Contains(p.Name))
                .ToList();

            //resolve displayed properties
            HashSet<string> allPropertyNames;
            if (DisplayedProperties != null)
            {
                allPropertyNames = DisplayedProperties;
            }
            else
            {
                allPropertyNames = new HashSet<string>(properties.Select(n => n.Name));
                if (dynamicProperties != null)
                {
                    foreach (var dynamicPropertyName in dynamicProperties.Select(n => n.Name))
                        allPropertyNames.Add(dynamicPropertyName);
                }
            }

            var displayFields = new List<DisplayField<T>>();
            foreach (var propertyName in allPropertyNames)
            {
                var property = properties.FirstOrDefault(n => n.Name == propertyName);
                if (property != null)
                {
                    //displayedProperties.Add(property); //todo: delete me
                    displayFields.Add(new DisplayField<T>(property));
                }
                else
                {
                    var dynamicProperty = DynamicProperties?.FirstOrDefault(n => n.Name == propertyName);
                    if (dynamicProperty != null)
                        displayFields.Add(new DisplayField<T>(dynamicProperty));
                }
            }

            //init the configurations
            var columnConfigurations = GetColumnConfigurations(displayFields.Select(n => n.Name));

            //create the worksheet
            var worksheet = package.Workbook.Worksheets.Add(WorksheetName);

            //Create table header
            {
                var col = headerFirstCol;
                foreach (var displayField in displayFields)
                {
                    var colConfig = columnConfigurations[displayField.Name];
                    var cell = worksheet.Cells[headerFirstRow, col];
                    cell.Value = string.IsNullOrEmpty(colConfig.Header.Text) ? displayField.DisplayName : colConfig.Header.Text;
                    colConfig.Header.SetStyle(cell.Style);

                    col++;
                }
            }

            //Add rows
            var row = dataFirstRow;
            foreach (var item in data)
            {
                var iCol = dataFirstCol;
                foreach (var displayField in displayFields)
                {
                    var cell = worksheet.Cells[row, iCol];
                    cell.Value = ApplyTextFormat(displayField.Name, displayField.GetValue(item));

                    iCol++;
                }
                row++;
            }
            
            //get bottom & right bounds
            var dataLastCol = dataFirstCol + displayFields.Count - 1;
            var dataLastRow = dataFirstRow + Math.Max(data.Count, 1) - 1; //make sure to have at least 1 data line (for table format)
            var tableRange = worksheet.Cells[headerFirstRow, headerFirstCol, dataLastRow, dataLastCol];
            
            WorksheetHelper.FormatAsTable(tableRange, TableStyle, WorksheetName, false);

            //apply configurations
            {
                var iCol = dataFirstCol;
                foreach (var displayField in displayFields)
                {
                    var colConfig = columnConfigurations[displayField.Name];
                    var columnRange = worksheet.Cells[dataFirstRow, iCol, dataLastRow, iCol];

                    //apply default number format
                    if (DefaultNumberFormats.ContainsKey(displayField.Type))
                        columnRange.Style.Numberformat.Format = DefaultNumberFormats[displayField.Type];

                    //apply number format
                    if (colConfig.Content.NumberFormat != null)
                        columnRange.Style.Numberformat.Format = colConfig.Content.NumberFormat;

                    //apply style
                    colConfig.Content.SetStyle(columnRange.Style);

                    if (colConfig.Width.HasValue)
                        worksheet.Column(iCol).Width = colConfig.Width.Value;
                    else if (AutoFitColumns)
                        worksheet.Column(iCol).AutoFit();


                    iCol++;
                }
            }
        
            //apply conditional styles
            {
                var iCol = dataFirstCol;
                foreach (var displayField in displayFields)
                {
                    if (ConditionalStyles.ContainsKey(displayField.Name))
                    {
                        var conditionalStyle = ConditionalStyles[displayField.Name];

                        var iRow = dataFirstRow;
                        foreach (var item in data)
                        {
                            var cell = worksheet.Cells[iRow, iCol];
                            conditionalStyle(item, cell.Style); //apply style on cell
                            iRow++;
                        }
                    }

                    iCol++;
                }
            }

            //apply conditional styles
            {
                var iCol = dataFirstCol;
                foreach (var displayField in displayFields)
                {
                    if (Formulas.ContainsKey(displayField.Name))
                    {
                        var formulaFormat = Formulas[displayField.Name];

                        var iRow = dataFirstRow;
                        foreach (var item in data)
                        {
                            var cell = worksheet.Cells[iRow, iCol];
                            var formula = formulaFormat(item, cell.Value); //apply style on cell
                            cell.Value = null;
                            cell.Formula = formula;
                            iRow++;
                        }
                    }

                    iCol++;
                }
            }
        
            //return the entire grid range
            return tableRange;
        }

        #endregion
    }
}
