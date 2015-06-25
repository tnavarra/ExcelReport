using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Reflect
{
    class Program
    {
        static void Main(string[] args)
        {
            List<PersonnelData> data = new List<PersonnelData>();
            data.Add(new PersonnelData { Location = "L1", Name = "N1" });
            data.Add(new PersonnelData { Location = "L2", Name = "N2" });
            data.Add(new PersonnelData { Location = "L3", Name = "N3" });
            data.Add(new PersonnelData { Location = "L4", Name = "N4" });
            data.Add(new PersonnelData { Location = "L5", Name = "N5" });
            data.Add(new PersonnelData { Location = "L6", Name = "N6" });


            ExcelReport<PersonnelData> report = new ExcelReport<PersonnelData>();
            report.DataRows = data;

            ExcelDataColumn<PersonnelData> column1 = new ExcelDataColumn<PersonnelData>();
            column1.SetColumnOrder(1).SetIsVisible(true).SetPropertyInfo("Name");


            report.Fields.Add(column1);

            foreach (var row in report.DataRows)
            {
                foreach (var column in report.Fields)
                {
                    Console.WriteLine(row.GetType().GetProperty(column.PropertyInfo.Name).GetValue(row));
                }
            }

            Console.ReadLine();
        }
    }
}
