using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace unvell.ReoGrid.IO.Additional
{
    static class ExcelDateTimeNumberFormatConverter
    {
        public static string ToReogridFormatString(string excelFormatString)
        {
            if (!excelFormatString.Contains("A/P") && !excelFormatString.Contains("AM/PM"))
            {
                excelFormatString = excelFormatString.Replace("hh", "HH");
            }
            return excelFormatString
                .Replace("yyyy/mm", "yyyy/MM")
                .Replace("mm/yy", "MM/yy")
                .Replace("mm/d", "MM/d")
                .Replace("m/d", "M/d")
                .Replace("d/mm", "d/MM")
                .Replace("d/m", "d/M")
                .Replace("aaaa", "dddd")
                .Replace("aaa", "ddd")
                .Replace("mmmm\\ yy", "MMMM\\ yy")
                .Replace("mmm\\-yy", "MMM\\-yy")
                ;
        }

        public static string ToExcelFormatString(string reogridFormatString)
        {
            return reogridFormatString;
        }
    }
}
