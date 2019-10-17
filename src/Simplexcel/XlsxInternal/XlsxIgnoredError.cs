﻿using System;
using System.Collections.Generic;
using Simplexcel.Cells;

namespace Simplexcel.XlsxInternal
{
    internal sealed class XlsxIgnoredError
    {
        private readonly IgnoredError _ignoredError;
        internal HashSet<CellAddress> Cells { get; }

        public XlsxIgnoredError()
        {
            Cells = new HashSet<CellAddress>();
            _ignoredError = new IgnoredError();
        }

        internal IgnoredError IgnoredError
        {
            get
            {
                // Note: This is a mutable reference, but changing it would be... bad.
                return _ignoredError;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                // And because the reference is mutable, this stores a copy so that modifications to the Worksheet don't break this
                _ignoredError.NumberStoredAsText = value.NumberStoredAsText;
                IgnoredErrorId = _ignoredError.GetHashCode();
            }
        }

        internal int IgnoredErrorId { get; private set; }

        internal string GetSqRef()
        {
            var ranges = CellAddressHelper.DetermineRanges(Cells);
            return string.Join(" ", ranges);
        }
    }
}
