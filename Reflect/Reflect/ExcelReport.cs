using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reflect
{
    public class ExcelReport<T> where T : class
    {
        public List<ExcelDataColumn<T>> Fields = new List<ExcelDataColumn<T>>();
        public IEnumerable<T> DataRows { get; set; }
    }
}
