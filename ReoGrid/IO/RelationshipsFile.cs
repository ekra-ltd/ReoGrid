using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;

namespace unvell.ReoGrid.IO.OpenXML.Schema
{
    public class RelationshipsFile
    {
        [XmlIgnore]
        internal Relationships _relationFile;
        [XmlIgnore]
        internal string _rsTarget;

        internal string GetAvailableRelationId()
        {
            if (_relationFile == null
                || _relationFile.relations == null
                || _relationFile.relations.Count == 0)
                return "rId1";

            int index = _relationFile.relations.Count + 1;
            string rId = null;

            while (_relationFile.relations.Any(s => s.id.Equals((rId = "rId" + index), StringComparison.CurrentCultureIgnoreCase)))
            {
                index++;
            }

            return rId;
        }

        internal string AddRelationship(string type, string targetFileName)
        {
            if (_relationFile == null)
            {
                _relationFile = new Relationships(_rsTarget);
            }

            if (_relationFile.relations == null)
            {
                _relationFile.relations = new List<Relationship>();
            }

            string rid = GetAvailableRelationId();

            _relationFile.relations.Add(new Relationship { id = rid, type = type, target = targetFileName });

            return rid;
        }
    }
}