using System.Xml.Serialization;

namespace unvell.ReoGrid.IO.Additional.Excel.FloatingObjects
{
    public partial class CT_LineSer
    {
        public CT_Marker_Chart2006 marker { get; set; }
    }


    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    public enum ST_MarkerStyle
    {
        circle,
        dash,
        diamond,
        dot,
        none,
        picture,
        plus,
        square,
        star,
        triangle,
        x,
        auto,
    }

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    public class CT_MarkerStyle
    {
        [XmlAttribute]
        public ST_MarkerStyle val { get; set; }
    }

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    public class CT_Marker_Chart2006
    {
        [XmlElement]
        public CT_MarkerStyle symbol { get; set; }
    }
}
