using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExporter.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Random r = new Random();

            var ds = new DataSet();
            var table1 = ds.Tables.Add("aaa");
            table1.Columns.Add("Id", typeof(long));
            table1.Columns.Add("Name", typeof(string));
            table1.Columns.Add("Date", typeof(DateTime));
            table1.Columns.Add("Guid", typeof(Guid));

            var row1 = table1.NewRow();
            row1["Id"] = 1;
            row1["Name"] = "Alper";
            row1["Date"] = new DateTime(2015,01,01);
            row1["Guid"] = Guid.NewGuid();
            table1.Rows.Add(row1);


            var row2 = table1.NewRow();
            row2["Id"] = 2;
            row2["Name"] = "Burcu";
            row2["Date"] = new DateTime(2015, 02, 02);
            row2["Guid"] = Guid.NewGuid();
            table1.Rows.Add(row2);

            var row3 = table1.NewRow();
            row3["Id"] = 3;
            row3["Name"] = "Yağmur";
            row3["Date"] = new DateTime(2015, 03, 03);
            row3["Guid"] = Guid.NewGuid();
            table1.Rows.Add(row3);

            var sec = "QWERTYUIOPĞÜASDFGHJKLŞİZXCVBNMÖÇ";

            for (int i = 0; i < 20; i++)
            {
                row3 = table1.NewRow();
                row3["Id"] = 3+i;
                row3["Name"] = string.Concat(Enumerable.Range(0,r.Next(4,8)).Select(x=>sec[r.Next(sec.Length)]));
                row3["Date"] = new DateTime(2015, r.Next(1,12), r.Next(1,28));
                row3["Guid"] = Guid.NewGuid();
                table1.Rows.Add(row3);
            }

            File.WriteAllBytes("Deneme.xlsx", ExcelExporter.GetExcelBytes("Deneme.xlsx", ds));
        }
    }
}
