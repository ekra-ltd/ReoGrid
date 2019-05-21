using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace unvell.ReoGrid.Formula
{
    /// <summary>
    /// 18.17.3 Error values
    /// </summary>
    enum ErrorValue
    {
        /// <summary>
        /// Intended to indicate when any number (including zero) or any error code is divided by zero
        /// </summary>
        Div0,
        /// <summary>
        /// Intended to indicate when a cell reference cannot be evaluated because the value for the cell has not been retrieved or calculated. [Note: This can happen when connected to an OLAP cube.end note]
        /// </summary>
        GettingData,
        /// <summary>
        /// Intended to indicate when a designated value is not available.
        /// </summary>
        /// <example>
        /// Some functions, such as SUMX2MY2, perform a series of operations on corresponding elements in two arrays. If those arrays do not have the same number of elements, then for some elements in the longer array, there are no corresponding elements in the shorter one; that is, one or more values in the shorter array are not available. 
        /// </example>
        NA,
        /// <summary>
        /// Intended to indicate when what looks like a name is used, but no such name has been defined.
        /// </summary>
        Name,
        /// <summary>
        /// Intended to indicate when two areas are required to intersect, but do not.
        /// </summary>
        Null,
        /// <summary>
        /// Intended to indicate when an argument to a function has a compatible type, but has a value that is outside the domain over which that function is defined. (This is known as a domain error.) 
        /// </summary>
        Num,
        /// <summary>
        /// Intended to indicate when a cell reference cannot be evaluated.
        /// </summary>
        Ref,
        /// <summary>
        /// Intended to indicate when an incompatible type argument is passed to a function, or an incompatible type operand is used with an operator.
        /// </summary>
        Value,
    }

    class ErrorConstant
    {
        private ErrorConstant(bool isUndefined, ErrorValue? value)
        {
            IsSetted = true;
            IsUndefined = isUndefined;
            ErrorValue = value;
        }

        public static ErrorConstant FromString(string formula)
        {
            if (NameConversation.ContainsKey(formula))
            {
                return new ErrorConstant(false, NameConversation[formula]);
            }
            return new ErrorConstant(true, null);
        }

        public bool IsSetted { get; }

        public bool IsUndefined { get; }

        public ErrorValue? ErrorValue { get; }

        #region Константы

        public const string Div0 = @"#DIV/0!";
        public const string GettingData = @"#GETTING_DATA";
        public const string NA = @"#N/A";
        public const string Name = @"#NAME?";
        public const string Null = @"#NULL!";
        public const string Num = @"#NUM!";
        public const string Ref = @"#REF!";
        public const string Value = @"#VALUE!";

        #endregion


        #region Приватные статические поля

        private static Dictionary<string, ErrorValue> NameConversation = new Dictionary<string, ErrorValue>
        {
            {Div0, Formula.ErrorValue.Div0},
            {GettingData, Formula.ErrorValue.GettingData},
            {NA, Formula.ErrorValue.NA},
            {Name, Formula.ErrorValue.Name},
            {Null, Formula.ErrorValue.Null},
            {Num, Formula.ErrorValue.Num},
            {Ref, Formula.ErrorValue.Ref},
            {Value, Formula.ErrorValue.Value}
        };

        #endregion
    }
}
