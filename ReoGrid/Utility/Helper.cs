using System.Text.RegularExpressions;
using unvell.ReoGrid.Formula;

namespace unvell.ReoGrid.Utility
{
    internal static class Helper
    {
        #region Конструктор

        static Helper()
        {
            SpPositionRegex = new Regex("^((('(?<escapedworksheetname>(([^']+)|(''))+)')|(?<worksheetname>[^!']+))!)?(?<cellId>\\$?[A-Z]+\\$?[0-9]+(:\\$?[A-Z]+\\$?[0-9]+)?)$");
        }

        #endregion

        #region Публичные методы

        public static SheetRangePosition ParseAsSheetRangePosition(IWorkbook workbook, string address)
        {
            var match = SpPositionRegex.Match(address);
            if (match.Success)
            {
                var escaped = match.Groups["escapedworksheetname"].Value;
                var name = match.Groups["worksheetname"].Value;
                var cell = match.Groups["cellId"].Value;

                if (string.IsNullOrEmpty(name))
                    name = escaped;

                if (!string.IsNullOrEmpty(name))
                {
                    if (RangePosition.IsValidAddress(cell))
                        return new SheetRangePosition(workbook?.GetWorksheetByName(name), new RangePosition(cell));
                }
            }
            return null;
        }

        #endregion

        #region Поля

        private static readonly Regex SpPositionRegex;

        #endregion
    }
}