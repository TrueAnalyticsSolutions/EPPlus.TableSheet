using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.TableSheet.Contracts.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace EPPlus.TableSheet
{
    /// <summary>
    /// Abstract representation of a structured table derived from the provided type (<typeparamref name="T"/>).
    /// </summary>
    /// <typeparam name="T">Reference to the type for which the structured table should contain many element.s</typeparam>
    public abstract partial class ExcelTableSheet<T>
    {
        /// <summary>
        /// Display text for the <see cref="ExcelWorksheet"/> that the structured table will be inserted.
        /// </summary>
        public abstract string WorksheetName { get; set; }

        /// <summary>
        /// Collection of <see cref="ExcelTableSheetColumn"/>s that make up the structured table columns for this table.
        /// </summary>
        public ICollection<ExcelTableSheetColumn> Properties { get; protected set; } = new List<ExcelTableSheetColumn>();

        /// <summary>
        /// Adds a new <see cref="ExcelTableSheetColumn"/> to <see cref="Properties"/>.
        /// </summary>
        /// <typeparam name="TProp"><inheritdoc cref="ExcelTableSheetColumn.Create" path="/typeparam[@name='TProp']"/></typeparam>
        /// <param name="key"><inheritdoc cref="ExcelTableSheetColumn.Create" path="/param[@name='key']"/></param>
        /// <param name="label"><inheritdoc cref="ExcelTableSheetColumn.Create" path="/param[@name='label']"/></param>
        /// <param name="format"><inheritdoc cref="ExcelTableSheetColumn.Create" path="/param[@name='format']"/></param>
        /// <returns></returns>
        public Guid AddProperty<TProp>(Expression<Func<T, TProp>> key, string label = "", string format = "")
        {
            Type propType = typeof(TProp);
            if (string.IsNullOrEmpty(label))
            {
                label = propType.Name;
            }

            Func<ExcelTableSheetColumn, string, bool> containsLabel = (p, l) => p.Label.Equals(l, StringComparison.OrdinalIgnoreCase);
            Func<string, bool> PropertiesContains = (l) => Properties.Any(o => containsLabel(o, l));

            if (PropertiesContains(label))
            {
                int nameIteration = 0;
                bool nameUnique = false;
                do
                {
                    nameIteration++;
                    if (!PropertiesContains($"{label}_{nameIteration}"))
                    {
                        nameUnique = true;
                    }
                } while (!nameUnique);
                label = $"{label}_{nameIteration}";
            }

            var property = ExcelTableSheetColumn.Create<TProp>(key, label, format);
            Properties.Add(property);
            return property.Guid;
        }

        /// <summary>
        /// Constructs a new <see cref="ExcelWorksheet"/> in the <paramref name="package"/> and inserts a structured, formatted table based on the provided source type (<typeparamref name="T"/>).
        /// </summary>
        /// <param name="package">Reference to the <see cref="ExcelPackage"/> that contains the <see cref="ExcelWorkbook"/> the <see cref="ExcelWorksheet"/> should be added to.</param>
        /// <param name="source">Collection of source objects (<typeparamref name="T"/>) to populate the structured table.</param>
        /// <param name="table">Outputs a reference to the structured table that was added to the return <see cref="ExcelWorksheet"/>.</param>
        /// <returns>Reference to the <see cref="ExcelWorksheet"/> that was added to the <paramref name="package"/>.</returns>
        public virtual ExcelWorksheet BuildTableSheet(ExcelPackage package, IEnumerable<T> source, out ExcelTable table)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(WorksheetName);

            table = BuildTable(worksheet, 1, source);
            // Auto-Fit Columns
            for (int i = 0; i < Properties.Count; i++)
            {
                try
                {
                    worksheet.Column(i + 1).AutoFit();
                }
                catch (Exception ex)
                {
                    // Maybe do something?
                }
            }
            return worksheet;
        }

        /// <summary>
        /// Constructs a new <see cref="ExcelTable"/> in the <paramref name="worksheet"/> based on the provided source type (<typeparamref name="T"/>).
        /// </summary>
        /// <param name="worksheet">Reference to the <see cref="ExcelWorksheet"/> for which the <see cref="ExcelTable"/> will be added.</param>
        /// <param name="startRow">The starting row for which the structured table should be inserted (1-based).</param>
        /// <param name="source"><inheritdoc cref="BuildTableSheet" path="/param[@name='source']"/></param>
        /// <returns></returns>
        protected virtual ExcelTable BuildTable(ExcelWorksheet worksheet, int startRow, IEnumerable<T> source)
        {
            int lastHeaderRow = BuildHeaderRow(worksheet, startRow);

            int lastBodyRow = BuildTableBody(worksheet, lastHeaderRow + 1, source);

            ExcelAddressBase tableRange = new ExcelAddressBase(lastHeaderRow, 1, lastBodyRow, Properties.Count);
            ExcelTable table = worksheet.Tables.Add(tableRange, WorksheetName.Replace(' ', '_'));
            return table;
        }

        /// <summary>
        /// Constructs the columns in the <paramref name="worksheet"/> for the formatted table.
        /// </summary>
        /// <param name="worksheet">Reference to the worksheet to add the table header.</param>
        /// <param name="startRow">Reference to the row for which to start adding the header columns.</param>
        /// <returns>Reference to which row the headers end. This may change if there are multiple rows of headers, ultimately the last row should be used for the table headers.</returns>
        protected virtual int BuildHeaderRow(ExcelWorksheet worksheet, int startRow)
        {
            if (startRow < 1) throw new IndexOutOfRangeException("Row index must be greater than or equal to 1.");

            int columnIndex = 1;
            foreach (var property in Properties)
            {
                // Add Header Label
                var headerCell = worksheet.Cells[startRow, columnIndex];
                headerCell.Value = property.Label;
                // Set Data Format?
                if (!string.IsNullOrEmpty(property.Format))
                {
                    worksheet.Column(columnIndex).Style.Numberformat.Format = property.Format;
                }

                columnIndex++;
            }
            return startRow;
        }

        /// <summary>
        /// Constructs the contents of the <see cref="ExcelTable"/> based on the provided source type (<typeparamref name="T"/>).
        /// </summary>
        /// <param name="worksheet">Reference to the worksheet to add the table data.</param>
        /// <param name="startRow">Reference to the row for which to start adding the data.</param>
        /// <param name="source"><inheritdoc cref="BuildTableSheet" path="/param[@name='source']"/></param>
        /// <returns></returns>
        protected virtual int BuildTableBody(ExcelWorksheet worksheet, int startRow, IEnumerable<T> source)
        {
            if (source.Count() < 1) return startRow;

            int row = startRow;
            foreach (var item in source)
            {
                int column = 1;
                foreach (var property in Properties)
                {
                    var valueCell = worksheet.Cells[row, column];
                    object value = null;
                    try
                    {
                        value = property.Getter.DynamicInvoke(new object[] { item });
                    }
                    catch (FailedToGetExcelTableValueException fex)
                    {
                        valueCell.AddComment($"Error: {fex}", "RevolutionSystem");
                    }
                    catch (TargetInvocationException tex)
                    {
                        value = string.Empty;
                    }
                    catch (Exception ex)
                    {
                        valueCell.AddComment($"Error: {ex}", "RevolutionSystem");
                    }
                    if (value != null && value != DBNull.Value)
                    {
                        if (property.PropertyType.IsEnum && value != null)
                        {
                            value = Enum.GetName(property.PropertyType, value);
                        }
                        valueCell.Value = value;
                    }
                    else
                    {
                        valueCell.Value = string.Empty;
                    }
                    column++;
                }
                row++;
            }

            return (row--); // Go back one because of for loop
        }
    }
}
