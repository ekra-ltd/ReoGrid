using System;
using System.Collections.Generic;
using unvell.ReoGrid.Drawing;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.IO.OpenXML;
using unvell.ReoGrid.Rendering;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid.IO.Additional.Excel.FloatingObjects
{
    class ImageObjectExporter: DrawingObjectExporterBase
    {
        #region DrawingObjectExporterBase

        public override bool CanExport(IDrawingObject exportObject, ExportOptions options)
        {
            if (exportObject is ImageObject image)
            {
                return image.Image != null;
            }
            return false;
        }

        public override void Export(Document doc, OpenXML.Schema.Worksheet sheet, OpenXML.Schema.Drawing drawing, Worksheet rgSheet, IDrawingObject exportObject, ExportOptions options)
        {
            if (CanExport(exportObject, options))
            {
                WriteImage(doc, sheet, drawing, rgSheet, exportObject as ImageObject);
            }
            else
            {
                throw new ArgumentException("", nameof(exportObject));
            }
        }

        #endregion

        private static void WriteImage(
            Document doc,
            OpenXML.Schema.Worksheet sheet,
            OpenXML.Schema.Drawing drawing,
            Worksheet rgSheet,
            Drawing.ImageObject image)
        {
            if (drawing.twoCellAnchors == null)
            {
                drawing.twoCellAnchors = new List<CT_TwoCellAnchor>();
            }

            string typeName = image.GetFriendlyTypeName();

            drawing._typeObjectCount.TryGetValue(typeName, out var typeObjCount);
            typeObjCount++;

            drawing._typeObjectCount[typeName] = typeObjCount;

            var twoCellAnchor = new CT_TwoCellAnchor
            {
                from = CreateCellAnchorByLocation_FormicrosoftXsd(rgSheet, image.Location),
                to = CreateCellAnchorByLocation_FormicrosoftXsd(rgSheet, new Point(image.Right, image.Bottom)),

                Item = new CT_Picture
                {
                    nvPicPr = new CT_PictureNonVisual
                    {
                        cNvPr = new CT_NonVisualDrawingProps
                        {
                            id = (uint)drawing._drawingObjectCount++,
                            name = typeName + " " + typeObjCount,
                        },

                        cNvPicPr = new CT_NonVisualPictureProperties
                        {
                            picLocks = new CT_PictureLocking(),
                        }
                    },

                    blipFill = new CT_BlipFillProperties
                    {
                        blip = doc.AddMediaImage_ForMicrosoftXsd(sheet, drawing, image),

                        stretch = new CT_StretchInfoProperties
                        {
                            fillRect = new CT_RelativeRect(),
                        },
                    },

                    spPr = CreateShapeProperty(image),
                },

                clientData = new CT_AnchorClientData(),
            };

            drawing.twoCellAnchors.Add(twoCellAnchor);
        }

       

        private static CT_ShapeProperties CreateShapeProperty(Drawing.DrawingObject image)
        {
            var dpi = PlatformUtility.GetDPI();

            return new CT_ShapeProperties
            {
                xfrm = new CT_Transform2D
                {
                    off = new CT_Point2D
                    {
                        x = MeasureToolkit.PixelToEMU(image.X, dpi),
                        y = MeasureToolkit.PixelToEMU(image.Y, dpi),
                    },
                    ext = new CT_PositiveSize2D
                    {
                        cx = MeasureToolkit.PixelToEMU(image.Right, dpi),
                        cy = MeasureToolkit.PixelToEMU(image.Bottom, dpi),
                    },
                },

                prstGeom = new CT_PresetGeometry2D
                {
                    prst = ST_ShapeType.rect,
                    avLst = new CT_GeomGuide[0],
                },
            };
        }

       
    }
}
