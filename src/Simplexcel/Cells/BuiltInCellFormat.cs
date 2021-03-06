﻿// ReSharper disable InconsistentNaming
namespace Simplexcel.Cells
{
    /// <summary>
    /// Excel Built-In Cell formats
    /// </summary>
    public static class BuiltInCellFormat
    {
        /// <summary>
        /// General
        /// </summary>
        public const string General = "General";

        /// <summary>
        /// 0
        /// </summary>
        public const string NumberNoDecimalPlaces = "0";

        /// <summary>
        /// 0.00
        /// </summary>
        public const string NumberTwoDecimalPlaces = "0.00";

        /// <summary>
        /// 0%
        /// </summary>
        public const string PercentNoDecimalPlaces = "0%";

        /// <summary>
        /// 0.00%
        /// </summary>
        public const string PercentTwoDecimalPlaces = "0.00%";

        /// <summary>
        /// @
        /// </summary>
        public const string Text = "@";

        /// <summary>
        /// m/d/yyyy h:mm
        /// </summary>
        public const string DateAndTime = "d/m/yyyy h:mm";

        /// <summary>
        /// m/d/yyyy
        /// </summary>
        public const string DateOnly = "d/m/yyyy";

        /// <summary>
        /// h:mm
        /// </summary>
        public const string TimeOnly = "h:mm";

        /// <summary>
        /// accounting
        /// </summary>
        public const string Accounting = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
    }
}
