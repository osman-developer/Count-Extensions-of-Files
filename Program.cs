using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
namespace LinqSamples
{
    public class Program {
        public static void Main () {
           
            var di = new DirectoryInfo (@"D:\Programming");
            var extensionCounts = di.EnumerateFiles ("*.*", SearchOption.AllDirectories)
                .GroupBy (x => x.Extension)
                .Select (g => new { Extension = g.Key, Count = g.Count () })
                .ToList ();

            using (var workbook = new XLWorkbook ()) {
                int i = 1;
                var worksheet = workbook.Worksheets.Add ("Sample Sheet");

                foreach (var group in extensionCounts) {
                    Console.WriteLine ("There are {0} files with extension {1}", group.Count,
                        group.Extension);

                    worksheet.Cell ("A" + i).Value = group.Extension;
                    worksheet.Cell ("B" + i).Value = group.Count;
                    i++;
                }

                workbook.SaveAs (@"D:\CountExtension.xlsx");
            }
        }
    }
}