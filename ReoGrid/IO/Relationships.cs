using System.Collections.Generic;
using System.Xml.Serialization;

namespace unvell.ReoGrid.IO.OpenXML.Schema
{
    [XmlRoot("Relationships", Namespace = OpenXMLNamespaces.Relationships)]
    public class Relationships
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns = new XmlSerializerNamespaces(
            new System.Xml.XmlQualifiedName[] {
                new System.Xml.XmlQualifiedName(string.Empty, OpenXMLNamespaces.Relationships),
            });

        [XmlElement("Relationship")]
        public List<Relationship> relations;

        [XmlIgnore]
        internal string _xmlTarget;

        public Relationships() { }

        internal Relationships(string _rsTarget)
        {
            _xmlTarget = _rsTarget;
        }
    }
}