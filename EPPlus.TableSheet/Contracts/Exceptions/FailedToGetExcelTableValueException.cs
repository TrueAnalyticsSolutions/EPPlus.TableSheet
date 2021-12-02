using System;

namespace OfficeOpenXml.TableSheet.Contracts.Exceptions
{
    /// <summary>
    /// Custom exception to be thrown when the <see cref="EPPlus.TableSheet.ExcelTableSheet{T}.ExcelTableSheetColumn.Getter"/> fails to get the value from the source type.
    /// </summary>
    public class FailedToGetExcelTableValueException : Exception
    {
        /// <summary>
        /// Constructs a new instance of <see cref="FailedToGetExcelTableValueException"/>.
        /// </summary>
        /// <param name="property">Reference to the property name that caused the failure.</param>
        /// <param name="message">Custom error message.</param>
        /// <param name="innerException">Exception that was thrown as a result of the attempted get method.</param>
        public FailedToGetExcelTableValueException(string property, string message, Exception innerException) : base($"{message}\r\nProperty: {property}", innerException) { }
    }
}
