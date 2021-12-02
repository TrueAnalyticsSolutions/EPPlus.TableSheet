namespace EPPlus.TableSheet
{
    /// <summary>
    /// Container of common and custom Excel cell formats
    /// </summary>
    public static class ExcelTableColumnFormats
    {   
        /// <summary>
        /// The standard Text format in Excel.
        /// </summary>
        public const string TEXT = "@";

        /// <summary>
        /// The standard Percent format in Excel.
        /// </summary>
        public const string PERCENT = "0.00%";

        /// <summary>
        /// The standard (US) Long Date format in Excel.
        /// </summary>
        public const string DATE_TIME = "m/d/yyyy h:mm:ss.ms";

        /// <summary>
        /// The standard Time format in Excel.
        /// </summary>
        public const string TIME = "h:mm:ss.ms";

        /// <summary>
        /// The standard (US) Short Date format in Excel.
        /// </summary>
        public const string DATE = "m/d/yyyy";

        /// <summary>
        /// A custom format for representing <see cref="System.TimeSpan"/> or other duration data types.
        /// </summary>
        public const string DURATION = "[hh]:mm:ss.ms";

        /// <summary>
        /// The standard Number format in Excel.
        /// </summary>
        public const string NUMBER = "0";

        /// <summary>
        /// The standard Number variant format in Excel that displays two decimal places.
        /// </summary>
        public const string NUMBER_TWO_DECIMAL_PLACES = "0.00";
    }
}
