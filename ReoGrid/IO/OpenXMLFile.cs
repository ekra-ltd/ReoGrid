using System.Xml.Serialization;

namespace unvell.ReoGrid.IO.OpenXML.Schema
{
    public class OpenXMLFile: RelationshipsFile
    {
        [XmlIgnore]
        internal string _resId;
        [XmlIgnore]
        internal string _xmlTarget;
        [XmlIgnore]
        internal string _path;
    }
}