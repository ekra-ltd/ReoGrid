using unvell.ReoGrid.Drawing;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.IO.OpenXML;
using unvell.ReoGrid.IO.OpenXML.Schema;
using unvell.ReoGrid.Rendering;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid.IO.Additional.Excel.FloatingObjects
{
    abstract class DrawingObjectExporterBase
    {
        public abstract bool CanExport(IDrawingObject exportObject, ExportOptions options);

        public abstract void Export(
            Document doc,
            OpenXML.Schema.Worksheet sheet,
            OpenXML.Schema.Drawing drawing,
            Worksheet rgSheet,
            IDrawingObject exportObject,
            ExportOptions options);

        protected static CellAnchor CreateCellAnchorByLocation(Worksheet rgSheet, Point p)
        {
            var dpi = PlatformUtility.GetDPI();

            int row = 0, col = 0;

            rgSheet.FindColumnByPosition(p.X, out col);
            rgSheet.FindRowByPosition(p.Y, out row);

            var colHeader = rgSheet.RetrieveColumnHeader(col);
            int colOffset = MeasureToolkit.PixelToEMU(p.X - colHeader.Left, dpi);

            var rowHeader = rgSheet.RetrieveRowHeader(row);
            int rowOffset = MeasureToolkit.PixelToEMU(p.Y - rowHeader.Top, dpi);

            return new CellAnchor
            {
                col = col,
                colOff = colOffset,
                row = row,
                rowOff = rowOffset
            };
        }

        protected static CT_Marker CreateCellAnchorByLocation_FormicrosoftXsd(Worksheet rgSheet, Point p)
        {
            var dpi = PlatformUtility.GetDPI();

            int row = 0, col = 0;

            rgSheet.FindColumnByPosition(p.X, out col);
            rgSheet.FindRowByPosition(p.Y, out row);

            var colHeader = rgSheet.RetrieveColumnHeader(col);
            int colOffset = MeasureToolkit.PixelToEMU(p.X - colHeader.Left, dpi);

            var rowHeader = rgSheet.RetrieveRowHeader(row);
            int rowOffset = MeasureToolkit.PixelToEMU(p.Y - rowHeader.Top, dpi);

            return new CT_Marker
            {
                col = col,
                colOff = colOffset,
                row = row,
                rowOff = rowOffset
            };
        }
    }
}
