﻿using System;
using Simplexcel.XlsxInternal;

namespace Simplexcel.Cells
{
    /// <summary>
    /// A cell inside a Worksheet
    /// </summary>
    public sealed class Cell
    {
        internal XlsxCellStyle XlsxCellStyle { get; }

        /// <summary>
        /// Create a new Cell of the given <see cref="CellType"/>.
        /// You can also implicitly create a cell from a string or number.
        /// </summary>
        /// <param name="cellType"></param>
        public Cell(CellType cellType) : this(cellType, null, BuiltInCellFormat.General) { }

        /// <summary>
        /// Create a new Cell of the given <see cref="CellType"/>, with the given value and format. For some common formats, see <see cref="BuiltInCellFormat"/>.
        /// You can also implicitly create a cell from a string or number.
        /// </summary>
        /// <param name="type"> </param>
        /// <param name="value"> </param>
        /// <param name="format"> </param>
        public Cell(CellType type, object value, string format)
        {
            XlsxCellStyle = new XlsxCellStyle
            {
                Format = format
            };

            Value = value;
            CellType = type;
            IgnoredErrors = new IgnoredError();
        }

        /// <summary>
        /// The Excel Format for the cell, see <see cref="BuiltInCellFormat"/>
        /// </summary>
        public string Format
        {
            get { return XlsxCellStyle.Format; }
            set { XlsxCellStyle.Format = value; }
        }

        /// <summary>
        /// The border around the cell
        /// </summary>
        public CellBorder Border
        {
            get { return XlsxCellStyle.Border; }
            set { XlsxCellStyle.Border = value; }
        }

        /// <summary>
        /// The name of the Font (Default: Calibri)
        /// </summary>
        public string FontName
        {
            get { return XlsxCellStyle.Font.Name; }
            set { XlsxCellStyle.Font.Name = value; }
        }

        /// <summary>
        /// The Size of the Font (Default: 11)
        /// </summary>
        public int FontSize
        {
            get { return XlsxCellStyle.Font.Size; }
            set { XlsxCellStyle.Font.Size = value; }
        }

        /// <summary>
        /// Should the text be bold?
        /// </summary>
        public bool Bold
        {
            get { return XlsxCellStyle.Font.Bold; }
            set { XlsxCellStyle.Font.Bold = value; }
        }

        /// <summary>
        /// Should the text be italic?
        /// </summary>
        public bool Italic
        {
            get { return XlsxCellStyle.Font.Italic; }
            set { XlsxCellStyle.Font.Italic = value; }
        }

        /// <summary>
        /// Should the text be underlined?
        /// </summary>
        public bool Underline
        {
            get { return XlsxCellStyle.Font.Underline; }
            set { XlsxCellStyle.Font.Underline = value; }
        }

        /// <summary>
        /// The font color.
        /// </summary>
        public Color TextColor
        {
            get { return XlsxCellStyle.Font.TextColor; }
            set { XlsxCellStyle.Font.TextColor = value; }
        }

        /// <summary>
        /// The interior/fill color.
        /// </summary>
        public PatternFill Fill
        {
            get { return XlsxCellStyle.Fill; }
            set { XlsxCellStyle.Fill = value; }
        }

        /// <summary>
        /// The Horizontal Alignment of content within a Cell
        /// </summary>
        public HorizontalAlign HorizontalAlignment
        {
            get { return XlsxCellStyle.HorizontalAlignment; }
            set { XlsxCellStyle.HorizontalAlignment = value; }
        }

        /// <summary>
        /// The Vertical Alignment of content within this Cell
        /// </summary>
        public VerticalAlign VerticalAlignment
        {
            get { return XlsxCellStyle.VerticalAlignment; }
            set { XlsxCellStyle.VerticalAlignment = value; }
        }

        /// <summary>
        /// Sets the Text to be Wrapped to new line
        /// </summary>
        public bool WrapText
        {
            get => XlsxCellStyle.WrapText;
            set => XlsxCellStyle.WrapText = value;
        }
        /// <summary>
        /// Any errors that are ignored in this Cell
        /// </summary>
        public IgnoredError IgnoredErrors { get; }

        /// <summary>
        /// The Type of the cell.
        /// </summary>
        public CellType CellType { get; }

        /// <summary>
        /// The Content of the cell.
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Should this cell be a Hyperlink to something?
        /// </summary>
        public string Hyperlink { get; set; }

        /// <summary>
        /// Create a new <see cref="Cell"/> that includes a Formula (e.g., SUM(A1:A5)). Do not include the initial = sign!
        /// </summary>
        /// <param name="formula">The formula, without the initial = sign (so "SUM(A1:A5)", not "=SUM(A1:A5)")</param>
        /// <returns></returns>
        public static Cell Formula(string formula)
        {
            return new Cell(CellType.Formula, formula, BuiltInCellFormat.General);
        }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Text from a string.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(string value)
        {
            return new Cell(CellType.Text, value, BuiltInCellFormat.Text);
        }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Number (formatted without decimal places) from an integer.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(long value)
        {
            return new Cell(CellType.Number, Convert.ToDecimal(value), BuiltInCellFormat.NumberNoDecimalPlaces);
        }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Number (formatted with 2 decimal places) places from a decimal.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(Decimal value)
        {
            return new Cell(CellType.Number, value, BuiltInCellFormat.NumberTwoDecimalPlaces);
        }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Number (formatted with 2 decimal places) places from a double.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(double value)
        {
            return new Cell(CellType.Number, Convert.ToDecimal(value), BuiltInCellFormat.NumberTwoDecimalPlaces);
        }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Date, formatted as DateAndTime.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(DateTime value)
        {
            return new Cell(CellType.Date, value, BuiltInCellFormat.DateAndTime);
        }

        /// <summary>
        /// Create a Cell from the given object, trying to determine the best cell type/format.
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static Cell FromObject(object val)
        {
            Cell cell;
            if (val is sbyte || val is short || val is int || val is long || val is byte || val is uint || val is ushort || val is ulong)
            {
                cell = new Cell(CellType.Number, Convert.ToDecimal(val), BuiltInCellFormat.NumberNoDecimalPlaces);
            }
            else if (val is float || val is double || val is decimal)
            {
                cell = new Cell(CellType.Number, Convert.ToDecimal(val), BuiltInCellFormat.General);
            }
            else if (val is DateTime)
            {
                cell = new Cell(CellType.Date, val, BuiltInCellFormat.DateAndTime);
            }
            else
            {
                cell = new Cell(CellType.Text, (val ?? String.Empty).ToString(), BuiltInCellFormat.Text);
            }
            return cell;
        }

        /// <summary>
        /// The largest positive number Excel can handle before <see cref="LargeNumberHandlingMode"/> applies
        /// </summary>
        public static decimal LargeNumberPositiveLimit => 99999999999m;

        /// <summary>
        /// The largest negative number Excel can handle before <see cref="LargeNumberHandlingMode"/> applies
        /// </summary>
        public static decimal LargeNumberNegativeLimit => -99999999999m;

        /// <summary>
        /// Check if the given number is so large that <see cref="LargeNumberHandlingMode"/> would apply to it
        /// </summary>
        /// <param name="number">The number to check</param>
        /// <returns></returns>
        public static bool IsLargeNumber(decimal? number) => number.HasValue && (number.Value < LargeNumberNegativeLimit || number.Value > LargeNumberPositiveLimit);
    }
}
