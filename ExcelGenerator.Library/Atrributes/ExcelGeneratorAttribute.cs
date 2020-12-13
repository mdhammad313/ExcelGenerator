using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelGenerator.Library.Atrributes
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public class ExcelGeneratorAttribute : Attribute
    {
        private readonly string _columnName;
        private readonly int _order;
        private readonly string _dateFormat;

        public ExcelGeneratorAttribute(string columnName = "", int order = 1, string dateFormat = null)
        {
            _columnName = columnName;
            _order = order;
            _dateFormat = dateFormat;
        }

        public string ColumnName { get { return _columnName; } private set { } }
        public int Order { get { return _order; } private set { } }
        public string DateFormat { get { return _dateFormat; } private set { } }
    }
}
