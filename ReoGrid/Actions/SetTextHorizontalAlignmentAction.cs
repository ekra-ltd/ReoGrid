namespace unvell.ReoGrid.Actions
{
    /// <summary>
    /// Set horizontal alignment for text cell
    /// </summary>
    public class SetTextHorizontalAlignmentAction : BaseWorksheetAction
    {
        /// <summary>
        /// Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName() => @"SetTextHorizontalAlignmentAction";

        RangePosition _position;
        ReoGridHorAlign _newValue;
        ReoGridHorAlign[,] _oldValues;

        /// <summary>
        /// Create instance for SetRowsHeightAction
        /// </summary>
        /// <param name="position">selected cells range</param>
        /// <param name="newValue">new value of horizontal alignment</param>
        public SetTextHorizontalAlignmentAction( RangePosition position,  ReoGridHorAlign newValue )
        {
            _position  = position;
            _newValue = newValue;
        }

        /// <summary>
        /// Do this action
        /// </summary>
        public override void Do()
        {
            if (_oldValues == null)
            {
                _oldValues = new ReoGridHorAlign[_position.Rows, _position.Cols];
                for (var i = _position.Row; i <= _position.EndRow; i++)
                    for (var j = _position.Col; j <= _position.EndCol; j++)
                    {
                        _oldValues[i - _position.Row, j - _position.Col] = Worksheet.Cells[i, j].Style.HAlign;
                        if (_oldValues[i - _position.Row, j - _position.Col] == ReoGridHorAlign.General)
                            _oldValues[i - _position.Row, j - _position.Col] = ReoGridHorAlign.Left;
                    }
            }

            for (var i = _position.Row; i <= _position.EndRow; i++)
                for (var j = _position.Col; j <= _position.EndCol; j++)
                {
                    Worksheet.Cells[i, j].Style.HAlign = _newValue;
                }
        }

        /// <summary>
        /// Undo this action
        /// </summary>
        public override void Undo()
        {
            for (var i = _position.Row; i <= _position.EndRow; i++)
                for (var j = _position.Col; j <= _position.EndCol; j++)
                {
                    Worksheet.Cells[i, j].Style.HAlign = _oldValues[i - _position.Row, j - _position.Col];
                }
        }
    }
}
