using System.Xml.Serialization;

namespace unvell.ReoGrid.IO.OpenXML.Schema
{
    public class SheetDrawingReference
    {
        [XmlAttribute("id", Namespace = OpenXMLNamespaces.R____________,
            Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string id;

#if DRAWING
        [XmlIgnore]
        internal Drawing _instance;
#endif // DRAWING

    }
}