using System.Xml.Serialization;

namespace unvell.ReoGrid.IO.OpenXML.Schema
{
    public class Relationship
    {
        [XmlAttribute("Id")]
        public string id;
        [XmlAttribute("Type")]
        public string type;
        [XmlAttribute("Target")]
        public string target;
    }
}