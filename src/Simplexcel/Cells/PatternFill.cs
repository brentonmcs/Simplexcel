using System;
// ReSharper disable NonReadonlyMemberInGetHashCode

namespace Simplexcel.Cells
{
    /// <summary>
    /// A Pattern Fill of a Cell.
    /// </summary>
    public class PatternFill : IEquatable<PatternFill>
    {
        // TODO: Add a Gradient Fill to this.
        // This isn't the structure in the XML, but that's how Excel Presents it, as a "Fill Effect"

        private bool _firstTimeSet;
        private Color? _bgColor;

        /// <summary>
        /// The type of Fill Pattern to use.
        /// Refer to the documentation for a list of each pattern.
        /// </summary>
        public PatternType PatternType { get; set; }

        /// <summary>
        /// The Background Color of the cell
        /// </summary>
        public Color? PatternColor { get; set; }

        /// <summary>
        /// The Background Color of the Fill.
        /// (No effect if <see cref="PatternType"/> is <see cref="PatternType.None"/>)
        /// </summary>
        public Color? BackgroundColor
        {
            get => _bgColor;
            set
            {
                _bgColor = value;

                // PatternType defaults to None, but the first time we set a background color,
                // set it to solid as the user likely wants the background color to show.
                // Further modifications won't change the PatternType, as this is now a deliberate setting
                if (_bgColor.HasValue && PatternType == PatternType.None && !_firstTimeSet)
                {
                    PatternType = PatternType.Solid;
                    _firstTimeSet = true;
                }
            }
        }

        /// <inheritdoc />
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(PatternFill)) return false;
            return Equals((PatternFill)obj);
        }

        /// <inheritdoc />
        public bool Equals(PatternFill other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(other.PatternType, PatternType)
                && Equals(other.PatternColor, PatternColor)
                && Equals(other.BackgroundColor, BackgroundColor);
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            unchecked
            {
                var result = PatternType.GetHashCode();
                if (PatternColor.HasValue)
                {
                    result = (result * 397) ^ PatternColor.GetHashCode();
                }
                if (BackgroundColor.HasValue)
                {
                    result = (result * 397) ^ BackgroundColor.GetHashCode();
                }
                return result;
            }
        }

        /// <summary>
        /// Handles equals
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator ==(PatternFill left, PatternFill right)
        {
            return Equals(left, right);
        }
        
        /// <summary>
        /// Handles not equal to
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator !=(PatternFill left, PatternFill right)
        {
            return !Equals(left, right);
        }
    }
}