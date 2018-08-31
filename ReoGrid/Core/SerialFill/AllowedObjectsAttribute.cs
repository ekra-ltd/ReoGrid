using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid.Core.SerialFill
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    abstract class AllowedObjectsAttribute: Attribute
    {
        public abstract bool IsAllow(object[] objects);
    }

    class SingleNumberAllowedObjectsAttribute : AllowedObjectsAttribute
    {
        public override bool IsAllow(object[] objects)
        {
            if (objects != null)
            {
                if (objects.Length == 1)
                {
                    double d;
                    return CellUtility.TryGetNumberData(objects[0], out d);
                }
            }
            return false;
        }
    }

    class NumberListAllowedObjectsAttribute : AllowedObjectsAttribute
    {
        public override bool IsAllow(object[] objects)
        {
            if (objects != null)
            {
                if (objects.Length > 1)
                {
                    return objects.All(obj =>
                    {
                        double d;
                        return CellUtility.TryGetNumberData(obj, out d);
                    });
                }
            }
            return false;
        }
    }

    class StringListAllowedObjectsAttributeAttribute : AllowedObjectsAttribute
    {
        public override bool IsAllow(object[] objects)
        {
            if (objects != null)
            {
                if (objects.Length >= 1)
                {
                    return objects.All(obj =>
                    {
                        return obj is string;
                    });
                }
            }
            return false;
        }
    }

    class DateTimeListAllowedObjectsAttribute : AllowedObjectsAttribute
    {
        public override bool IsAllow(object[] objects)
        {
            if (objects != null)
            {
                if (objects.Length >= 1)
                {
                    return objects.All(obj =>
                    {
                        return obj == null || obj is DateTime ;
                    });
                }
            }
            return false;
        }
    }
}
