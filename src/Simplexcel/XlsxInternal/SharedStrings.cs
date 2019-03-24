﻿using System;
using System.Linq;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// The Shared Strings table in an Excel document
    /// </summary>
    internal class SharedStrings
    {
        private readonly Dictionary<string, int> _sharedStrings = new Dictionary<string, int>(StringComparer.Ordinal);

        /// <summary>
        /// The number of Unique Strings
        /// </summary>
        internal int UniqueCount { get { return _sharedStrings.Count; } }
        /// <summary>
        /// The number of total Strings (incl. duplicate uses of the same string)
        /// </summary>
        internal int Count { get; private set; }

        /// <summary>
        /// Add a string to the shared strings table and return the index of that string
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        internal int GetStringIndex(string input)
        {
            // NULL is treated as an empty string
            if (input == null)
            {
                input = string.Empty;
            }

            Count++;
            if (!_sharedStrings.ContainsKey(input))
            {
                _sharedStrings[input] = _sharedStrings.Count;
            }
            return _sharedStrings[input];
        }

        /// <summary>
        /// Create the XmlFile for the Package.
        /// </summary>
        /// <returns></returns>
        internal XmlFile ToXmlFile()
        {
            var file = new XmlFile();
            file.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
            file.Path = "xl/sharedStrings.xml";

            var sst = new XDocument(new XElement(Namespaces.x + "sst",
                    new XAttribute("xmlns", Namespaces.x),
                    new XAttribute("count", Count),
                    new XAttribute("uniqueCount", UniqueCount)
                ));

            foreach (var kvp in _sharedStrings.OrderBy(k => k.Value))
            {
                var tElem = new XElement(Namespaces.x + "t");
                tElem.Value = kvp.Key;

                var siElem = new XElement(Namespaces.x + "si", tElem);
                sst.Root.Add(siElem);
            }

            file.Content = sst;

            return file;
        }

        internal Relationship ToRelationship(RelationshipCounter relationshipCounter)
        {
            return new Relationship(relationshipCounter)
                {
                    Target = ToXmlFile(),
                    Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
                };
        }
    }
}
