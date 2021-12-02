using System;
using System.Linq.Expressions;

namespace EPPlus.TableSheet
{
    public abstract partial class ExcelTableSheet<T>
    {
        /// <summary>
        /// Instance of a structured table column with references to property information of the parent type (<typeparamref name="T"/>).
        /// </summary>
        public class ExcelTableSheetColumn
        {
            /// <summary>
            /// Reference to the property this column derives from the parent type (<typeparamref name="T"/>).
            /// </summary>
            public Type PropertyType { get; }

            /// <summary>
            /// Display text for the header of the table column.
            /// </summary>
            public string Label { get; set; }

            /// <summary>
            /// Method to get the property value from the parent type (<typeparamref name="T"/>)
            /// </summary>
            public Delegate Getter { get; private set; }

            /// <summary>
            /// Display format for the values of this table column.
            /// </summary>
            public string Format { get; set; } = string.Empty;

            public Guid Guid { get; private set; } = System.Guid.NewGuid();

            /// <summary>
            /// Creates a new instance of an Excel structured table column.
            /// </summary>
            /// <param name="propertyType"><inheritdoc cref="PropertyType" path="/summary"/></param>
            /// <param name="valueGetter"><inheritdoc cref="Getter" path="/summary"/></param>
            /// <param name="label"><inheritdoc cref="Label" path="/summary"/></param>
            /// <param name="format"><inheritdoc cref="Format" path="/summary"/></param>
            public ExcelTableSheetColumn(Type propertyType, Delegate valueGetter, string label = "", string format = "")
            {
                PropertyType = propertyType;

                if (string.IsNullOrEmpty(label))
                {
                    label = PropertyType.Name;
                }
                Label = label;
                if (!string.IsNullOrEmpty(format))
                {
                    Format = format;
                }
                else if (PropertyType == typeof(DateTime))
                {
                    Format = ExcelTableColumnFormats.DATE_TIME;
                }
                else if (PropertyType == typeof(TimeSpan))
                {
                    Format = ExcelTableColumnFormats.DURATION;
                }
                Getter = valueGetter;
            }

            /// <summary>
            /// Creates a new <see cref="ExcelTableSheetColumn"/> from the provided expression.
            /// </summary>
            /// <typeparam name="TProp">Generic reference to the property type.</typeparam>
            /// <param name="key"><inheritdoc cref="Getter" path="/summary"/></param>
            /// <param name="label"><inheritdoc cref="Label" path="/summary"/></param>
            /// <param name="format"><inheritdoc cref="Format" path="/summary"/></param>
            /// <returns></returns>
            public static ExcelTableSheetColumn Create<TProp>(Expression<Func<T, TProp>> key, string label = "", string format = "")
              => new ExcelTableSheetColumn(typeof(TProp), key.Compile(), label, format);
        }
    }
}
