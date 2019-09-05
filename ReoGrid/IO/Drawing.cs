using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using unvell.ReoGrid.IO.Additional.Excel.FloatingObjects;

namespace unvell.ReoGrid.IO.OpenXML.Schema
{
    [XmlRoot("wsDr", Namespace = OpenXMLNamespaces.XDR__________)]
    public class Drawing : OpenXMLFile
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns = new XmlSerializerNamespaces(
            new System.Xml.XmlQualifiedName[]
            {
                new System.Xml.XmlQualifiedName("xdr", OpenXMLNamespaces.XDR__________),
                new System.Xml.XmlQualifiedName("a", OpenXMLNamespaces.Drawing______),
            });

        [XmlElement("twoCellAnchor")]
        // public List<TwoCellAnchor> twoCellAnchors;
        public List<CT_TwoCellAnchor> twoCellAnchors;

        [XmlIgnore]
        internal int _drawingObjectCount = 2;

        [XmlIgnore]
        internal Dictionary<string, int> _typeObjectCount;

        [Obsolete]
        [XmlIgnore]
        internal List<Blip> _images;

        [XmlIgnore]
        internal List<CT_Blip> _images_ForMicrosoftXsd;

        [XmlIgnore]
        internal CT_ChartSpace _chartSpace;
    }
}