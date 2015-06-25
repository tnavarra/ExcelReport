using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Reflect
{
    public class ExcelDataColumn<T> where T : class
    {
        public PropertyInfo PropertyInfo { get; private set; }
        public bool IsVisible { get; private set; }
        public int ColumnOrder { get; private set; }

        public ExcelDataColumn<T> SetPropertyInfo(string name)
        {
            PropertyInfo = typeof(T).GetProperty(name);
            return this;
        }

        public ExcelDataColumn<T> SetIsVisible(bool isVisible)
        {
            IsVisible = isVisible;
            return this;
        }

        public ExcelDataColumn<T> SetColumnOrder(int order)
        {
            ColumnOrder = order;
            return this;
        }
    }
}
