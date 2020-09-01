using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml;
using unvell.ReoGrid.Core;
using unvell.ReoGrid.Core.Worksheet.Additional;
using unvell.ReoGrid.Formula;
using unvell.ReoGrid.Graphics;
using E2006 = unvell.ReoGrid.IO.OpenXML.Schema;
using E2009 = unvell.ReoGrid.IO.OpenXML.Schema.Excel2009;

namespace unvell.ReoGrid
{
    public partial class Worksheet
    {
        public List<ConditionalFormat> ConditionalFormats { get; set; }

        private readonly Dictionary<ConditionalFormat, List<ReferenceRange>> cfFormulaRanges =
            new Dictionary<ConditionalFormat, List<ReferenceRange>>();

        public void TryAddConditionalFormats()
        {
            if (ConditionalFormats != null)
                foreach (var format in ConditionalFormats)
                {
                    try
                    {
                        TryAddConditionalFormat(format);
                    }
                    catch
                    {
                        // молчим
                    }
                }
        }

        private static readonly List<ConditionalFormatAdder> ConditionalFormatAdders = new List<ConditionalFormatAdder>(
            new ConditionalFormatAdder[]
            {
                new ConditionalFormatExpressionAdder(),
                new ConditionalFormatCellIsAdder(),
            });

        private void TryAddConditionalFormat(ConditionalFormat format)
        {
            var position = new RangePosition(format.Sqref.Text).StartPos;

            foreach (var rule in format.Rules)
            {
                if (rule.DifferentialFormat == null) continue;
                ConditionalFormatAdders.FirstOrDefault(a => a.ExpressionType == rule.Type)?.Add(new AddConditionalFormatInfo
                {
                    Worksheet = this,
                    Rule = rule,
                    Format = format,
                    Position = position
                });
            }
        }

        public void RecalcConditionalFormats()
        {
            if (ConditionalFormats != null)
                foreach (var format in ConditionalFormats)
                {
                    RecalcConditionalFormat(format);
                }
        }

        internal void RecalcConditionalFormats(Cell cell)
        {
            foreach (var sheet in Workbook.Worksheets)
            {
                foreach (var range in sheet.cfFormulaRanges)
                {
                    bool applyed = false;
                    foreach (var reference in range.Value)
                    {
                        if (reference.Cells.Contains(cell))
                        {
                            sheet.RecalcConditionalFormat(range.Key);
                            applyed = true;
                        }
                        if (applyed) break;
                    }
                }
            }
        }

        internal void RecalcConditionalFormats(RangePosition position)
        {
        }

        private void RecalcConditionalFormat(ConditionalFormat format)
        {
            var pos = new RangePosition(format.Sqref.Text);
            var cell = Cells[pos.StartPos];

            for (int r = 0; r < pos.Rows; r++)
            for (int c = 0; c < pos.Cols; c++)
            {
                RestoredefaultDifferencialFormat(pos, c, r);
            }

            foreach (var rule in format.Rules.OrderByDescending(r => r.Priority))
            {
                var dxFormat = rule.DifferentialFormat;
                if (dxFormat == null) continue;

                var cfAdder = ConditionalFormatAdders.FirstOrDefault(a => a.ExpressionType == rule.Type);
                if (cfAdder != null)
                {
                    for (int r = 0; r < pos.Rows; r++)
                    for (int c = 0; c < pos.Cols; c++)
                    {
                        if (cfAdder.CanApplyFormat(Cells[cell.Row + r, cell.Column + c], c, r, rule))
                        {
                            ApplyDifferencialFormat(pos, c, r, dxFormat);
                        }
                    }
                }
            }
        }

        private void RestoredefaultDifferencialFormat(RangePosition position, int columnOffset, int rowOffset)
        {
            ApplyDifferencialFormat(position, columnOffset, rowOffset, null);
        }

        private void ApplyDifferencialFormat(RangePosition position, int columnOffset, int rowOffset, DifferentialFormat format)
        {
            var cell = Cells[position.Row + rowOffset, position.Col + columnOffset];
            cell.ApplyDifferencialFormat(format);
        }

        public void ResetConditionalFormatting()
        {
            foreach (var cell in CfCells)
            {
                cell?.ClearCfFormat();
            }
        }

        public void MarkAsCfCell(Cell cell)
        {
            if(!CfCells.Contains(cell))
                CfCells.Add(cell);
        }

        private  HashSet<Cell> CfCells = new HashSet<Cell>();

        #region Вспомогательные методы

        private class AddConditionalFormatInfo
        {
            public Worksheet Worksheet { get; set; }
            public ConditionalFormatRule Rule { get; set; }
            public CellPosition Position { get; set; }
            public ConditionalFormat Format { get; set; }
        }

        private abstract class ConditionalFormatAdder
        {
            public abstract ConditionalFormatType ExpressionType { get; }
            public abstract void Add(AddConditionalFormatInfo info);

            /// <summary>
            /// Проверку условия срабатвания условного форматирования
            /// </summary>
            /// <param name="cell">правый верхний угол области условного форматирования (SqRef)</param>
            /// <param name="columnOffset">Смещение вправо относитильно правого верхнего угла</param>
            /// <param name="rowOffset">Смещение вниз относитильно правого верхнего угла</param>
            /// <param name="rule">Правило условного форматирования</param>
            /// <returns>true, если необходимо применить правило для ячейки, иначе - false</returns>
            public abstract bool CanApplyFormat(Cell cell, int columnOffset, int rowOffset, ConditionalFormatRule rule);

            protected void ExpandReferenceRanges(AddConditionalFormatInfo info, List<ReferenceRange> referencesRanges)
            {
                if (string.IsNullOrEmpty(info?.Format?.Sqref?.Text))
                    return;

                var adder = RangePosition.Empty;
                try
                {
                    adder = new RangePosition(info.Format.Sqref.Text);
                }
                catch
                {
                    return;
                }

                if(adder.Cols < 1) return;
                if (adder.Rows < 1) return;

                foreach (var range in referencesRanges.OfType<FormulaReferenceRange>())
                {
                    try
                    {
                        if (range?.FormulaCellPosition is null) continue;
                        if (range.FormulaCellPosition.ColumnProperty == PositionProperty.Relative)
                        {
                            range.Cols = adder.Cols - 1;
                        }
                        if (range.FormulaCellPosition.RowProperty == PositionProperty.Relative)
                        {
                            range.Rows = adder.Rows - 1;
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            /// <summary>
            /// Функция выполняет сдвиг ячеек, адрес которых задан относительным
            /// </summary>
            /// <param name="node">Исходное дерево формулы</param>
            /// <param name="columnOffset">смещение вправо</param>
            /// <param name="rowOffset">смещение влево</param>
            /// <returns></returns>
            protected static STNode ShiftNode(STNode node, int columnOffset, int rowOffset)
            {
                if (columnOffset == 0 && rowOffset == 0)
                    return node;
                else
                {
                    var clone = node.Clone() as STNode;
                    if (clone != null)
                    {
                        STNode.RecursivelyIterate(clone, stNode => ShiftSingleNode(stNode, columnOffset, rowOffset));
                    }
                    return clone;
                }
            }

            private static void ShiftSingleNode(STNode node, int columnOffset, int rowOffset)
            {
                if (node != null)
                {
                    if (node is STCellNode cellNode)
                    {

                        if (cellNode.Position != null)
                        {
                            var colShift = cellNode.Position.ColumnProperty == PositionProperty.Relative;
                            var rowShift = cellNode.Position.RowProperty == PositionProperty.Relative;
                            var newPosition = new CellPosition(cellNode.Position.Row + (rowShift ? rowOffset : 0), cellNode.Position.Col + (colShift ? columnOffset : 0));
                            cellNode.Position = newPosition;
                        }
                    }
                    else if (node is STRangeNode tangeNode)
                    {
                        
                    }
                }
            }
        }

        private class ConditionalFormatExpressionAdder : ConditionalFormatAdder
        {
            public override ConditionalFormatType ExpressionType => ConditionalFormatType.Expression;

            public override void Add(AddConditionalFormatInfo info)
            {
                if (info.Rule.Formula.Count > 0)
                {
                    try
                    {
                        var node = Parser.Parse(info.Worksheet.Workbook, info.Worksheet.Cells[info.Position], info.Rule.Formula[0].Value);
                        if (info.Worksheet.cfFormulaRanges.TryGetValue(info.Format, out var referencesRanges))
                        {
                            referencesRanges.Clear();
                        }
                        else
                        {
                            referencesRanges = info.Worksheet.cfFormulaRanges[info.Format] = new List<ReferenceRange>();
                        }

                        if (node != null)
                        {
                            try
                            {
                                IterateToAddReference(info.Worksheet.Cells[info.Position], node, referencesRanges, true);
                                ExpandReferenceRanges(info, referencesRanges);
                                info.Rule.Formula[0].FormulaTree = node;
                            }
                            catch (CircularReferenceException)
                            {
                                info.Rule.Formula[0].FormulaTree = null;
                                throw;
                            }
                        }
                    }
                    catch (Exception e)
                    {
#if DEBUG
                        MessageBox.Show($"При разборе формулы условного форматирования диапазона '{info.Format.Sqref.Text}' '{info.Rule.Formula[0].Value}' выдано исключение: '{e.Message}'");
#endif
                        Debug.WriteLine(e);
                        throw;
                    }
                }
            }

            public override bool CanApplyFormat(Cell cell, int columnOffset, int rowOffset, ConditionalFormatRule rule)
            {
                var evaluatedValue = Evaluator.Evaluate(cell, ShiftNode(rule.Formula[0].FormulaTree, columnOffset, rowOffset));
                return evaluatedValue.type == FormulaValueType.Boolean && evaluatedValue.value as bool? == true;
            }
        }

        private class ConditionalFormatCellIsAdder : ConditionalFormatAdder
        {
            public override ConditionalFormatType ExpressionType => ConditionalFormatType.CellIs;

            public override void Add(AddConditionalFormatInfo info)
            {
                foreach (var formulaItem in info.Rule.Formula )
                {
                    try
                    {
                        var node = Parser.Parse(info.Worksheet.Workbook, info.Worksheet.Cells[info.Position], formulaItem.Value);
                        if (info.Worksheet.cfFormulaRanges.TryGetValue(info.Format, out var referencesRanges))
                        {
                            referencesRanges.Clear();
                        }
                        else
                        {
                            referencesRanges = info.Worksheet.cfFormulaRanges[info.Format] = new List<ReferenceRange>();
                        }

                        if (node != null)
                        {
                            try
                            {
                                IterateToAddReference(info.Worksheet.Cells[info.Position], node, referencesRanges, false);
                                ExpandReferenceRanges(info, referencesRanges);
                                if (!referencesRanges.Any(r =>  r.Contains(info.Position)))
                                    referencesRanges.Add(new ReferenceRange(info.Worksheet, info.Position));
                                formulaItem.FormulaTree = node;
                            }
                            catch (CircularReferenceException)
                            {
                                formulaItem.FormulaTree = null;
                                throw;
                            }
                        }
                    }
                    catch (Exception e)
                    {
#if DEBUG
                        MessageBox.Show($"При разборе формулы условного форматирования диапазона '{info.Format.Sqref.Text}' '{formulaItem.Value}' выдано исключение: '{e.Message}'");
#endif
                        Debug.WriteLine(e);
                        throw;
                    }
                }
            }

            public override bool CanApplyFormat(Cell cell, int columnOffset, int rowOffset, ConditionalFormatRule rule)
            {
                List<FormulaValue> values = new List<FormulaValue>();
                foreach (var formulaItem in rule.Formula)
                {
                    values.Add(Evaluator.Evaluate(cell, ShiftNode(formulaItem.FormulaTree, columnOffset, rowOffset)));
                }
                // var evaluatedValue = Evaluator.Evaluate(cell, rule.Formula[0].FormulaTree);
                return CheckCondition(rule.Operator, cell?.Data, values);
            }

            private static bool CheckCondition(ConditionalFormattingOperator? op, object cellvalue, List<FormulaValue> evaluatedValues)
            {
                if (op is null) return false;
                if (cellvalue is null) return false;
                if (evaluatedValues is null) return false;
                if (!evaluatedValues.Any()) return false;
                if (evaluatedValues.Any(i => i.value is null)) return false;

                switch (op)
                {
                    case ConditionalFormattingOperator.LessThan: return IsLessThenAsDouble(cellvalue, evaluatedValues[0].value);
                    case ConditionalFormattingOperator.LessThanOrEqual: return IsLessThanOrEqualAsDouble(cellvalue, evaluatedValues[0].value);
                    case ConditionalFormattingOperator.Equal: return IsEqualAsDouble(cellvalue, evaluatedValues[0].value);
                    case ConditionalFormattingOperator.NotEqual: return IsNotEqualAsDouble(cellvalue, evaluatedValues[0].value);
                    case ConditionalFormattingOperator.GreaterThanOrEqual: return IsGreaterThanOrEqualAsDouble(cellvalue, evaluatedValues[0].value);
                    case ConditionalFormattingOperator.GreaterThan: return IsGreaterThanAsDouble(cellvalue, evaluatedValues[0].value);
                    case ConditionalFormattingOperator.Between:
                        if (evaluatedValues.Count >= 2)
                        {
                            return IsBetweenAsDouble(cellvalue, evaluatedValues[0].value, evaluatedValues[1].value);
                        }
                        break;
                    case ConditionalFormattingOperator.NotBetween:
                        if (evaluatedValues.Count >= 2)
                        {
                            return IsNotBetweenAsDouble(cellvalue, evaluatedValues[0].value, evaluatedValues[1].value);
                        }
                        break;
                }
                return false;
            }

            private static bool IsLessThenAsDouble(object a, object b)
            {
                if (AsDouble(a, out var dA) && AsDouble(b, out var dB))
                {
                    return dA < dB;
                }
                return false;
            }

            private static bool IsLessThanOrEqualAsDouble(object a, object b)
            {
                if (AsDouble(a, out var dA) && AsDouble(b, out var dB))
                {
                    return dA <= dB;
                }
                return false;
            }

            private static bool IsGreaterThanOrEqualAsDouble(object a, object b)
            {
                if (AsDouble(a, out var dA) && AsDouble(b, out var dB))
                {
                    return dA >= dB;
                }
                return false;
            }

            private static bool IsGreaterThanAsDouble(object a, object b)
            {
                if (AsDouble(a, out var dA) && AsDouble(b, out var dB))
                {
                    return dA > dB;
                }
                return false;
            }

            private static double Epsilon => 0.00001;

            private static bool IsEqualAsDouble(object a, object b)
            {
                if (AsDouble(a, out var dA) && AsDouble(b, out var dB))
                {
                    return Math.Abs(dA - dB) < Epsilon;
                }
                return false;
            }

            private static bool IsNotEqualAsDouble(object a, object b)
            {
                if (AsDouble(a, out var dA) && AsDouble(b, out var dB))
                {
                    return !(Math.Abs(dA - dB) < Epsilon);
                }
                return false;
            }

            private static bool IsBetweenAsDouble(object a, object b, object c)
            {
                return IsLessThenAsDouble(b, c)
                    ? IsGreaterThanOrEqualAsDouble(a, b) && IsLessThanOrEqualAsDouble(a, c)
                    : IsLessThanOrEqualAsDouble(a, b) && IsGreaterThanOrEqualAsDouble(a, c);
            }

            private static bool IsNotBetweenAsDouble(object a, object b, object c)
            {
                return IsLessThenAsDouble(b, c)
                    ? IsLessThenAsDouble(a, b) || IsGreaterThanAsDouble(a, c)
                    : IsGreaterThanAsDouble(a, b) || IsLessThenAsDouble(a, c);
            }

            private static bool AsDouble(object o, out double d)
            {
                if (o is double doubleValue)
                {
                    d = doubleValue;
                    return true;
                }
                d = double.NaN;
                return false;
            }

        }

        #endregion
    }

    public partial class Cell
    {
        private CfSaveStyle _cfSaveStyle { get; set; }

        private CfApplyStyle _cfOverrideStyle { get; set; }

        internal void ClearCfFormat()
        {
            if (_cfSaveStyle != null)
            {
                Style.BackColor = _cfSaveStyle.BackColor;
                Style.TextColor = _cfSaveStyle.TextColor;
                Style.Bold = _cfSaveStyle.Bold;
                Style.Italic = _cfSaveStyle.Italic;
            }
        }


        internal void ApplyDifferencialFormat(DifferentialFormat format)
        {
            Worksheet?.MarkAsCfCell(this);

            if (_cfSaveStyle == null)
                _cfSaveStyle = new CfSaveStyle
                {
                    BackColor = Style.BackColor,
                    TextColor = Style.TextColor,
                    Bold = Style.Bold,
                    Italic = Style.Italic,
                };

            if (format is null)
            {
                _cfOverrideStyle = null;
            }
            else
            {
                var bkColor = format.Fill?.PatternFill?.BackgroundColor;
                var txtColor = format.Font?.Color;
                var sBold = format.Font?.Bold?.Value ?? false;
                var sItalic = format.Font?.Italic?.Value ?? false;

                SolidColor? sBackColor = bkColor?.RgbColorValue != null
                    ? (SolidColor?) new SolidColor(
                        bkColor.RgbColorValue.Value[0]
                        , bkColor.RgbColorValue.Value[1]
                        , bkColor.RgbColorValue.Value[2]
                        , bkColor.RgbColorValue.Value[3])
                    : null;
                SolidColor? sTextColor = txtColor?.RgbColorValue != null
                    ? (SolidColor?) new SolidColor(
                        txtColor.RgbColorValue.Value[0]
                        , txtColor.RgbColorValue.Value[1]
                        , txtColor.RgbColorValue.Value[2]
                        , txtColor.RgbColorValue.Value[3])
                    : null;

                _cfOverrideStyle = new CfApplyStyle
                {
                    BackColor = sBackColor,
                    TextColor = sTextColor,
                    Bold = sBold,
                    Italic = sItalic,
                };
            }

            if (_cfOverrideStyle is null)
            {
                Style.BackColor = _cfSaveStyle.BackColor;
                Style.TextColor = _cfSaveStyle.TextColor;
                Style.Bold = _cfSaveStyle.Bold;
                Style.Italic = _cfSaveStyle.Italic;
                if (true == InnerStyle?.HasStyle(PlainStyleFlag.TextColor))
                {
                    InnerStyle.TextColor = Style.TextColor;
                }
                Worksheet?.UpdateCellFont(this, UpdateFontReason.TextColorChanged);
            }
            else
            {
                if(_cfOverrideStyle.BackColor.HasValue)
                    Style.BackColor = _cfOverrideStyle.BackColor.Value;
                if (_cfOverrideStyle.TextColor.HasValue)
                {
                    Style.TextColor = _cfOverrideStyle.TextColor.Value;
                    if (true == InnerStyle?.HasStyle(PlainStyleFlag.TextColor))
                    {
                        InnerStyle.TextColor = Style.TextColor;
                    }
                    Worksheet?.UpdateCellFont(this, UpdateFontReason.TextColorChanged);
                }
                Style.Bold = _cfOverrideStyle.Bold;
                Style.Italic = _cfOverrideStyle.Italic;
            }

            // if (_cfOverrideStyle.BackColor != null)
            // {
            //     if (_cfOverrideStyle is null)
            //         Style.BackColor = _cfSaveStyle.BackColor;
            //     else
            //         Style.BackColor = _cfOverrideStyle.BackColor.Value;
            // }
            // if (_cfOverrideStyle.TextColor != null)
            // {
            //     if (_cfOverrideStyle is null)
            //         Style.TextColor = _cfSaveStyle.TextColor;
            //     else
            //         Style.TextColor = _cfOverrideStyle.TextColor.Value;
            //     if (true == InnerStyle?.HasStyle(PlainStyleFlag.TextColor))
            //     {
            //         InnerStyle.TextColor = Style.TextColor;
            //     }
            //     Worksheet?.UpdateCellFont(this, UpdateFontReason.TextColorChanged);
            // }
        }

        [Serializable]
        private class CfSaveStyle
        {
            public SolidColor BackColor { get; set; }
            public SolidColor TextColor { get; set; }
            public bool Bold { get; set; }
            public bool Italic { get; set; }
        }

        [Serializable]
        private class CfApplyStyle
        {
            public SolidColor? BackColor { get; set; }
            public SolidColor? TextColor { get; set; }
            public bool Bold { get; set; }
            public bool Italic { get; set; }
        }

    }

}

namespace unvell.ReoGrid.Core.Worksheet.Additional
{
    internal static class ConditionalFormatHelper
    {
        public static List<ConditionalFormat> From2006(E2006.CT_ConditionalFormatting[] formattings, E2006.Stylesheet stylesheet, IO.OpenXML.Document doc)
        {
            var result = new List<ConditionalFormat>();
            foreach (var formatting in formattings)
            {
                try
                {
                    result.Add(From2006(formatting, stylesheet, doc));
                }
                catch 
                {
                    //skip
                }
            }
            return result;
        }

        public static List<ConditionalFormat> From2009(E2009.CT_ConditionalFormattings formattings, IO.OpenXML.Document doc)
        {
            //throw new NotImplementedException();
            List<ConditionalFormat> result = new List<ConditionalFormat>();
            foreach (var formatting in formattings.conditionalFormatting)
            {
                result.Add(From2009(formatting, doc));
            }
            return result;
        }

        public static E2009.CT_ConditionalFormattings ToExcel2009(List<ConditionalFormat> formats)
        {
            E2009.CT_ConditionalFormattings result = null;
            if (formats != null && formats.Count > 0)
            {
                result = new E2009.CT_ConditionalFormattings
                {
                };
                var list = new List<E2009.CT_ConditionalFormatting>();
                foreach (var format in formats)
                {
                    var item = ToExcel2009(format);
                    if (item != null)
                        list.Add(item);
                }
                if (list.Count > 0)
                    result.conditionalFormatting = list.ToArray();
            }
            return result;
        }

        #region 2006->RG

        private static ConditionalFormat From2006(E2006.CT_ConditionalFormatting formatting, E2006.Stylesheet stylesheet, IO.OpenXML.Document doc)
        {
            var result = new ConditionalFormat();
            result.Pivot = formatting.pivot;
            foreach (var cfRule in formatting.cfRule)
            {
                result.Rules.Add(From2006(cfRule, stylesheet, doc));
            }
            result.Sqref = From2006(formatting.sqref);
            return result;
        }

        private static Sqref From2006(string[] value)
        {
            Sqref result = null;
            if (value != null)
            {
                result = new Sqref
                {
                    Text = string.Concat(value),
                    Adjust = null,
                    Adjusted = null,
                    Edited = null,
                    Split = null,
                };
            }
            return result;
        }

        private static ConditionalFormatRule From2006(E2006.CT_CfRule rule, E2006.Stylesheet stylesheet, IO.OpenXML.Document doc)
        {
            if (rule != null)
            {
                string extId = null;
                if (rule.extLst != null)
                {
                    // выполняем поиск по расширениям, ищем id, если находим - значит это продублированный в секции 2009 условное форматирование
                    foreach (var ext in rule.extLst.ext)
                    {
                        if (ext.Any != null && ext.uri == "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}")
                        {
                            var enumerator = ext.Any.GetEnumerator();
                            while(enumerator.MoveNext())
                            {
                                var data = enumerator.Current;
                                XmlText text = data as XmlText;
                                if (text != null)
                                {
                                    extId = text.Value;
                                }
                            }
                        }
                    }
                }
                var result = new ConditionalFormatRule
                {
                    IconSet = From2006(rule.iconSet),
                    AboveAverage = rule.aboveAverage,
                    ActivePercent = null,
                    Bottom = rule.bottom,
                    ColorScale = From2006(rule.colorScale, doc),
                    DataBar = From2006(rule.dataBar, doc),
                    DifferentialFormat = From2006(rule.dxfId, rule.dxfIdSpecified, stylesheet),
                    EqualAverage = rule.equalAverage,
                    Operator = From2006(rule.@operator, rule.operatorSpecified),
                    Percent = rule.percent,
                    Priority = rule.priority,
                    Rank = rule.rankSpecified?(uint?) rule.rank:null,
                    SGuid = null,
                    StdDev = rule.stdDevSpecified?(int?)rule.stdDev:null,
                    StopIfTrue = rule.stopIfTrue,
                    Text = rule.text,
                    TimePeriod = From2006(rule.timePeriod, rule.timePeriodSpecified),
                    Type = From2006(rule.type, rule.typeSpecified),
                    Ext2009Id = extId,
                };
                if (rule.formula != null)
                {
                    foreach (var s in rule.formula)
                    {
                        result.Formula.Add(From2006(s));
                    }
                }
                return result;
            }
            return null;
        }

        private static IconSet From2006(E2006.CT_IconSet value)
        {
            IconSet result = null;
            if (value != null)
            {
                result = new IconSet
                {
                    Custom = null, //TODO перепроверить
                    Percent = value.percent,
                    IconSetType = From2006(value.iconSet),
                    Reverse = value.reverse,
                    ShowValues = value.showValue,
                };
                foreach (var cfvo in value.cfvo)
                {
                    result.CondittionalFormatValue.Add(From2006(cfvo));
                }
                //TODO перепроверить result.Cficon
            }
            return result;
            throw new NotImplementedException();
        }

        private static ColorScale From2006(E2006.CT_ColorScale value, IO.OpenXML.Document doc)
        {
            ColorScale result = null;
            if (value != null)
            {
                result = new ColorScale{};
                foreach (var cfvo in value.cfvo)
                {
                    result.CondittionalFormatValue.Add(From2006(cfvo));
                }
                foreach (var color in value.color)
                {
                    result.Color.Add(From2006(color, doc));
                }
            }
            return result;
        }

        private static DataBar From2006(E2006.CT_DataBar value, IO.OpenXML.Document doc)
        {
            DataBar result = null;
            if (value != null)
            {
                result = new DataBar
                {
                    AxisColor = null,
                    AxisPosition = null,
                    Border = null,
                    BorderColor = null,
                    Direction = null,
                    FillColor = From2006(value.color, doc),
                    Gradient = null,
                    MaxLength = value.maxLength,
                    MinLength = value.minLength,
                    NegativeBarBorderColorSameAsPositive = null,
                    NegativeBarColorSameAsPositive = null,
                    NegativeBorderColor = null,
                    NegativeFillColor = null,
                    ShowValue = value.showValue,
                };
                foreach (var cfvo in value.cfvo)
                {
                    result.CondittionalFormatValue.Add(From2006(cfvo));
                }
            }
            return result;
        }

        private static DifferentialFormat From2006(uint value, bool specified, E2006.Stylesheet stylesheet)
        {

            DifferentialFormat result = null;
            if (specified && value < stylesheet?.differentialFormats.Count)
            {
                var dxf = stylesheet.differentialFormats.list[(int)value];
                result = new DifferentialFormat
                {
                    Border = From2006(dxf.Boarder),
                    CellAlignment = From2006(dxf.Alignment),
                    CellProtection = From2006(dxf.Protection),
                    Fill = From2006(dxf.Fill),
                    Font = From2006(dxf.Font),
                    NumberFormat = From2006(dxf.NumFmt)
                };
            }
            return result;
        }

        private static FormulaItem From2006(string formula)
        {
            FormulaItem result = null;
            if (formula != null)
            {
                result = new FormulaItem
                {
                    Value = formula
                };
            }
            return result;
        }

        private static ConditionalFormattingOperator? From2006(E2006.ST_ConditionalFormattingOperator value, bool specified)
        {
            if (!specified) return null;
            //TODO на входе может и не быть значения
            if (value != null)
            {
                switch (value)
                {
                    case E2006.ST_ConditionalFormattingOperator.lessThan:
                        return ConditionalFormattingOperator.LessThan;
                    case E2006.ST_ConditionalFormattingOperator.lessThanOrEqual:
                        return ConditionalFormattingOperator.LessThanOrEqual;
                    case E2006.ST_ConditionalFormattingOperator.equal:
                        return ConditionalFormattingOperator.Equal;
                    case E2006.ST_ConditionalFormattingOperator.notEqual:
                        return ConditionalFormattingOperator.NotEqual;
                    case E2006.ST_ConditionalFormattingOperator.greaterThanOrEqual:
                        return ConditionalFormattingOperator.GreaterThanOrEqual;
                    case E2006.ST_ConditionalFormattingOperator.greaterThan:
                        return ConditionalFormattingOperator.GreaterThan;
                    case E2006.ST_ConditionalFormattingOperator.between:
                        return ConditionalFormattingOperator.Between;
                    case E2006.ST_ConditionalFormattingOperator.notBetween:
                        return ConditionalFormattingOperator.NotBetween;
                    case E2006.ST_ConditionalFormattingOperator.containsText:
                        return ConditionalFormattingOperator.ContainsText;
                    case E2006.ST_ConditionalFormattingOperator.notContains:
                        return ConditionalFormattingOperator.NotContains;
                    case E2006.ST_ConditionalFormattingOperator.beginsWith:
                        return ConditionalFormattingOperator.BeginsWith;
                    case E2006.ST_ConditionalFormattingOperator.endsWith:
                        return ConditionalFormattingOperator.EndsWith;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            return null;
        }

        private static TimePeriod? From2006(E2006.ST_TimePeriod value, bool specified)
        {
            if (!specified)
                return null;
            //TODO на входе может и не быть значения
            if (value != null)
            {
                switch (value)
                {
                    case E2006.ST_TimePeriod.today:
                        return TimePeriod.Today;
                    case E2006.ST_TimePeriod.yesterday:
                        return TimePeriod.Yesterday;
                    case E2006.ST_TimePeriod.tomorrow:
                        return TimePeriod.Tomorrow;
                    case E2006.ST_TimePeriod.last7Days:
                        return TimePeriod.Last7Days;
                    case E2006.ST_TimePeriod.thisMonth:
                        return TimePeriod.ThisMonth;
                    case E2006.ST_TimePeriod.lastMonth:
                        return TimePeriod.LastMonth;
                    case E2006.ST_TimePeriod.nextMonth:
                        return TimePeriod.NextMonth;
                    case E2006.ST_TimePeriod.thisWeek:
                        return TimePeriod.ThisWeek;
                    case E2006.ST_TimePeriod.lastWeek:
                        return TimePeriod.LastWeek;
                    case E2006.ST_TimePeriod.nextWeek:
                        return TimePeriod.NextWeek;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            return null;
        }

        private static ConditionalFormatType? From2006(E2006.ST_CfType value, bool specified)
        {
            if (!specified) return null;
            //TODO на входе может и не быть значения
            if (value != null)
            {
                switch (value)
                {
                    case E2006.ST_CfType.expression:
                        return ConditionalFormatType.Expression;
                    case E2006.ST_CfType.cellIs:
                        return ConditionalFormatType.CellIs;
                    case E2006.ST_CfType.colorScale:
                        return ConditionalFormatType.ColorScale;
                    case E2006.ST_CfType.dataBar:
                        return ConditionalFormatType.DataBar;
                    case E2006.ST_CfType.iconSet:
                        return ConditionalFormatType.IconSet;
                    case E2006.ST_CfType.top10:
                        return ConditionalFormatType.Top10;
                    case E2006.ST_CfType.uniqueValues:
                        return ConditionalFormatType.UniqueValues;
                    case E2006.ST_CfType.duplicateValues:
                        return ConditionalFormatType.DuplicateValues;
                    case E2006.ST_CfType.containsText:
                        return ConditionalFormatType.ContainsText;
                    case E2006.ST_CfType.notContainsText:
                        return ConditionalFormatType.NotContainsText;
                    case E2006.ST_CfType.beginsWith:
                        return ConditionalFormatType.BeginsWith;
                    case E2006.ST_CfType.endsWith:
                        return ConditionalFormatType.EndsWith;
                    case E2006.ST_CfType.containsBlanks:
                        return ConditionalFormatType.ContainsBlanks;
                    case E2006.ST_CfType.notContainsBlanks:
                        return ConditionalFormatType.NotContainsBlanks;
                    case E2006.ST_CfType.containsErrors:
                        return ConditionalFormatType.ContainsErrors;
                    case E2006.ST_CfType.notContainsErrors:
                        return ConditionalFormatType.NotContainsErrors;
                    case E2006.ST_CfType.timePeriod:
                        return ConditionalFormatType.TimePeriod;
                    case E2006.ST_CfType.aboveAverage:
                        return ConditionalFormatType.AboveAverage;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            return null;
        }

        private static IconSetType From2006(E2006.ST_IconSetType value)
        {
            // TODO Проверить может ли тип быть пустым
            //throw new NotImplementedException();
            switch (value)
            {
                case E2006.ST_IconSetType.Item3Arrows:
                    return IconSetType.IconSet_3Arrows;
                case E2006.ST_IconSetType.Item3ArrowsGray:
                    return IconSetType.IconSet_3ArrowsGray;
                case E2006.ST_IconSetType.Item3Flags:
                    return IconSetType.IconSet_3Flags;
                case E2006.ST_IconSetType.Item3TrafficLights1:
                    return IconSetType.IconSet_3TrafficLights1;
                case E2006.ST_IconSetType.Item3TrafficLights2:
                    return IconSetType.IconSet_3TrafficLights2;
                case E2006.ST_IconSetType.Item3Signs:
                    return IconSetType.IconSet_3Signs;
                case E2006.ST_IconSetType.Item3Symbols:
                    return IconSetType.IconSet_3Symbols;
                case E2006.ST_IconSetType.Item3Symbols2:
                    return IconSetType.IconSet_3Symbols2;
                case E2006.ST_IconSetType.Item4Arrows:
                    return IconSetType.IconSet_4Arrows;
                case E2006.ST_IconSetType.Item4ArrowsGray:
                    return IconSetType.IconSet_4ArrowsGray;
                case E2006.ST_IconSetType.Item4RedToBlack:
                    return IconSetType.IconSet_4RedToBlack;
                case E2006.ST_IconSetType.Item4Rating:
                    return IconSetType.IconSet_4Rating;
                case E2006.ST_IconSetType.Item4TrafficLights:
                    return IconSetType.IconSet_4TrafficLights;
                case E2006.ST_IconSetType.Item5Arrows:
                    return IconSetType.IconSet_5Arrows;
                case E2006.ST_IconSetType.Item5ArrowsGray:
                    return IconSetType.IconSet_5ArrowsGray;
                case E2006.ST_IconSetType.Item5Rating:
                    return IconSetType.IconSet_5Rating;
                case E2006.ST_IconSetType.Item5Quarters:
                    return IconSetType.IconSet_5Quarters;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static ConditionalFormatValueObject From2006(E2006.CT_Cfvo value)
        {
            ConditionalFormatValueObject result = null;
            if (value != null)
            {
                result = new ConditionalFormatValueObject
                {
                    Formula = From2006(value.val),
                    Gte = value.gte,
                    Type = From2006(value.type)
                };
                // TODO extList - проигнорирован
            }
            return result;
        }

        private static Color From2006(E2006.CT_Color value, IO.OpenXML.Document doc)
        {
            Color result = null;
            if (value != null)
            {
                byte[] rgb = value.rgb;
                if (value.themeSpecified && doc?.Themesheet?.elements?.clrScheme != null)
                {
                    uint theme = value.theme;
                    var s = doc.Themesheet?.elements?.clrScheme.GetElement(theme);
                    if (s?.srgbClr?.val != null)
                    {
                        if (s.srgbClr.val.Length == 3)
                        {
                            rgb = new[]
                            {
                                (byte) 255,
                                s.srgbClr.val[0],
                                s.srgbClr.val[1],
                                s.srgbClr.val[2],
                            };
                        }else if (s.srgbClr.val.Length == 4)
                        {
                            rgb = new[]
                            {
                                s.srgbClr.val[0],
                                s.srgbClr.val[1],
                                s.srgbClr.val[2],
                                s.srgbClr.val[3],
                            };
                        }
                    }
                    //if (TextFormatHelper.DecodeColor(s.srgbClr.val, out var color))
                    //{
                    //    rgb = new[] {color.A, color.R, color.G, color.B};
                    //}
                }
                result = new Color
                {
                    Automatic = value.autoSpecified ? (bool?) value.auto : null,
                    Indexed = value.indexedSpecified ? (uint?) value.indexed : null,
                    ThemeColor = value.themeSpecified ? (uint?) value.theme : null,
                    RgbColorValue = From2006(rgb),
                    TInt = value.tint,
                };
            }
            return result;
        }

        private static Border From2006(E2006.Border value)
        {
            Border result = new Border();
            if (value != null)
            {
                result = new Border
                {
                    Bottom = From2006(value.bottom),
                    Diagonal = From2006(value.diagonal),
                    DiagonalDown = null,        // TODO несовпадение имен наверное
                    DiagonalUp = null,
                    End = null,
                    Horizontal = null,
                    Outline = null,
                    Start = null,
                    Top = From2006(value._top),
                    Vertical = null,
                };
            }
            return result;
        }

        private static CellAlignment From2006(E2006.Alignment value)
        {
            CellAlignment result = null;
            if (value != null)
            {
                int indent;
                bool indetParsed = int.TryParse(value.indent, out indent);
                bool wrap;
                bool wrapParsed = bool.TryParse(value.wrapText, out wrap);
                result = new CellAlignment
                {
                    Horizontal = From2006(value._horAlign),
                    Vertical = From2006(value._verAlign),
                    Indent = null,
                    JustifyLastLine = null,
                    ReadingOrder = null,
                    RelativeIndent = (indetParsed ? (int?)indent: null),
                    ShrinkToFit = null,
                    TextRotation = value.textRotation,
                    WrapText = wrapParsed? wrap:false,
                };
            }
            return result;
        }

        private static CellProtection From2006(E2006.Protection value)
        {
            CellProtection result = null;
            if (value != null)
            {
                bool locked;
                bool lockedParsed = bool.TryParse(value.locked, out locked);
                result = new CellProtection
                {
                    Hidden = false,
                    Locked = lockedParsed ? locked : false,
                };
            }
            return result;
        }

        private static Fill From2006(E2006.Fill value)
        {
            Fill result = null;
            if (value != null)
            {
                result = new Fill
                {
                    GradientFill = null,
                    PatternFill = From2006(value.patternFill),
                };
            }
            return result;
        }

        private static Font From2006(E2006.Font value)
        {
            Font result = null;
            if (value != null)
            {
                bool strike = value.strikethrough?.val != null && value.strikethrough?.val != "0";
                bool bold = value.bold != null && value.bold.val != "0";
                bool italic = value.italic != null && value.italic.val != "0";
                result = new Font
                {
                    Bold = bold ? new BooleanProperty {Value = true} : null,
                    Color = From2006(value.color),
                    Outline = ToBooleanProperty(null),
                    Charset = null,
                    Condense = ToBooleanProperty(null),
                    Extend = ToBooleanProperty(null),
                    Family = From2006FontFamily(value.family),
                    FontScheme = null,
                    FontSize = From2006FontSize(value.size),
                    Italic = italic ? new BooleanProperty {Value = true} : null,
                    Name = From2006FontName(value.name),
                    Shadow = ToBooleanProperty(null),
                    Strike = strike ? new BooleanProperty {Value = true} : null,
                    Underline = From2006(value.underline),
                    VerticalAlign = null,
                };
            }
            return result;
        }

        private static NumberFormat From2006(E2006.NumberFormat value)
        {
            NumberFormat result = null;
            if (value != null)
            {
                result = new NumberFormat
                {
                    FormatCode = value.formatCode,
                    NumberFormatId = (uint)value.formatId,  // TODO что за formatId
                };
            }
            return result;
        }

        private static ConditionalFormatValueObjectType From2006(E2006.ST_CfvoType value)
        {
            //if (value != null)
            {
                switch (value)
                {
                    case E2006.ST_CfvoType.num:
                        return ConditionalFormatValueObjectType.Num;
                    case E2006.ST_CfvoType.percent:
                        return ConditionalFormatValueObjectType.Percent;
                    case E2006.ST_CfvoType.max:
                        return ConditionalFormatValueObjectType.Max;
                    case E2006.ST_CfvoType.min:
                        return ConditionalFormatValueObjectType.Min;
                    case E2006.ST_CfvoType.formula:
                        return ConditionalFormatValueObjectType.Formula;
                    case E2006.ST_CfvoType.percentile:
                        return ConditionalFormatValueObjectType.Percentile;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
        }

        private static Argb From2006(byte[] rgb)
        {
            //TODO а я в не в курсе как выполнять преобразование
            Argb result = new Argb();
            if (rgb != null)
            {
                Array.Copy(rgb, result.Value, Math.Min(result.Value.Length, rgb.Length));
            }
            return result;
        }

        private static BorderPr From2006(RangeBorderStyle value)
        {
            BorderPr result = null;
            if (value != null)
            {
                result = new BorderPr
                {
                    Color = From2006(value.Color),
                    Style = From2006(value.Style),
                };
            }
            return result;
        }

        private static BorderPr From2006(E2006.SideBorder value)
        {
            BorderPr result = null;
            if (value != null)
            {
                result = new BorderPr
                {
                    Color = From2006(value.color),
                    Style = From2006BorderStyle(value.style),   //TODO непонятно какая строка должна быть
                };
            }
            return result;
        }

        private static HorizontalAlignment From2006(ReoGridHorAlign value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case ReoGridHorAlign.General:
                        return HorizontalAlignment.General;
                    case ReoGridHorAlign.Left:
                        return HorizontalAlignment.Left;
                    case ReoGridHorAlign.Center:
                        return HorizontalAlignment.Center;;
                    case ReoGridHorAlign.Right:
                        return HorizontalAlignment.Right;
                    case ReoGridHorAlign.DistributedIndent:
                        return HorizontalAlignment.Distributed;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            // TODO Проверить может ли быть значение null
            throw new ArgumentNullException(nameof(value));
        }

        private static VerticalAlignment From2006(ReoGridVerAlign value)
        {
            if (value != null)

            {
                switch (value)
                {
                    case ReoGridVerAlign.General:
                        return VerticalAlignment.Distributed; //TODO нет general - может быть другой тип
                    case ReoGridVerAlign.Top:
                        return VerticalAlignment.Top;
                    case ReoGridVerAlign.Middle:
                        return VerticalAlignment.Center;
                    case ReoGridVerAlign.Bottom:
                        return VerticalAlignment.Bottom;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            // TODO Проверить может ли быть значение null
            throw new ArgumentNullException(nameof(value));
        }

        private static PatternFill From2006(E2006.PatternFill value)
        {
            PatternFill result = null;
            if (value != null)
            {
                result = new PatternFill
                {
                    BackgroundColor = From2006(value.backgroundColor),
                    ForegroundColor = From2006(value.foregroundColor),
                    PatternType = From2006PatternType(value.patternType),
                };
            }
            return result;
        }

        private static BooleanProperty ToBooleanProperty(object o)
        {
            return new BooleanProperty {Value = o != null};
        }

        private static Color From2006(E2006.ColorValue value)
        {
            Color result = null;
            if (value != null)
            {
                bool auto;
                bool autoParsed = bool.TryParse(value.auto, out auto);

                uint indexed;
                bool indexedPartsed = uint.TryParse(value.indexed, out indexed);

                uint themeColor;
                bool themeColorParsed = uint.TryParse(value.theme, out themeColor);

                double tint;
                bool timtParsed = double.TryParse(value.tint?.Replace(".", ","), out tint);

                result = new Color
                {
                    Automatic = autoParsed ? (bool?) auto : null,
                    Indexed = indexedPartsed ? (uint?) indexed : null,
                    ThemeColor = themeColorParsed? (uint?)themeColor: null,
                    TInt = timtParsed?(double?)tint:null,
                    RgbColorValue = From2006Argb(value.rgb),
                };
            }
            return result;
        }

        private static FontFamily From2006FontFamily(E2006.ElementValue<string> value)
        {
            FontFamily result = null;
            if (value != null)
            {
                result = new FontFamily
                {
                    Value = value.value
                };
            }
            return result;
        }

        private static FontSize From2006FontSize(E2006.ElementValue<string> value)
        {
            FontSize result = null;
            if (value != null)
            {
                double d;
                if (double.TryParse(value.value, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                {
                    result = new FontSize
                    {
                        Value = d,
                    };
                }
            }
            return result;
        }

        private static FontName From2006FontName(E2006.ElementValue<string> value)
        {

            FontName result = null;
            if (value != null)
            {
                result = new FontName
                {
                    Value = value.value
                };
            }
            return result;
        }

        private static UnderlineProperty From2006(E2006.Underline value)
        {
            if (value != null)
            {
                return new UnderlineProperty {Value = UnderlineValues.Single};
            }
            return new UnderlineProperty {Value = UnderlineValues.None};
        }

        private static Color From2006(SolidColor value)
        {
            Color result = null;
            if (value != null)
            {
                result = new Color
                {
                    Automatic = null,
                    Indexed = null,
                    ThemeColor = null,
                    TInt = value.IsTransparent ? 0.0f : 1.0f,
                    RgbColorValue = From2006Argb(value.A, value.R, value.G, value.B),
                };
            }
            return result;
        }

        private static BorderStyle? From2006(BorderLineStyle value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case BorderLineStyle.None:
                        return BorderStyle.None;
                    case BorderLineStyle.Solid:
                        return BorderStyle.Thin;    // TODO мне не известно соответствие
                    case BorderLineStyle.Dotted:
                        return BorderStyle.Dotted;
                    case BorderLineStyle.Dashed:
                        return  BorderStyle.Dashed;
                    case BorderLineStyle.DoubleLine:
                        return BorderStyle.Double;
                    case BorderLineStyle.Dashed2:
                        return BorderStyle.Dashed;
                    case BorderLineStyle.DashDot:
                        return BorderStyle.DashDot;
                    case BorderLineStyle.DashDotDot:
                        return BorderStyle.DashDotDot;
                    case BorderLineStyle.BoldDashDot:
                        return BorderStyle.DashDot;
                    case BorderLineStyle.BoldDashDotDot:
                        return BorderStyle.MediumDashDotDot;
                    case BorderLineStyle.BoldDashed:
                        return BorderStyle.Dashed;
                    case BorderLineStyle.BoldDotted:
                        return BorderStyle.Dotted;
                    case BorderLineStyle.BoldSolid:
                        return BorderStyle.Thin;
                    case BorderLineStyle.BoldSolidStrong:
                        return BorderStyle.Thin;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
                
            }
            return null;
        }

        private static BorderStyle? From2006BorderStyle(string value)
        {
            if (value != null)
            {
                BorderLineStyle style;
                if (Enum.TryParse(value, false, out style))
                {
                    return From2006(style);
                }
            }
            return null;
        }

        private static PatternType? From2006PatternType(string value)
        {
            if (value != null)
            {
                E2009.ST_PatternType style;
                if (Enum.TryParse(value, false, out style))
                {
                    switch (style)
                    {
                        case E2009.ST_PatternType.none:
                            return PatternType.None;
                        case E2009.ST_PatternType.solid:
                            return PatternType.Solid;
                        case E2009.ST_PatternType.mediumGray:
                            return PatternType.MediumGray;
                        case E2009.ST_PatternType.darkGray:
                            return PatternType.DarkGray;
                        case E2009.ST_PatternType.lightGray:
                            return  PatternType.LightGray;
                        case E2009.ST_PatternType.darkHorizontal:
                            return PatternType.DarkHorizontal;
                        case E2009.ST_PatternType.darkVertical:
                            return PatternType.DarkVertical;
                        case E2009.ST_PatternType.darkDown:
                            return PatternType.DarkDown;
                        case E2009.ST_PatternType.darkUp:
                            return PatternType.DarkUp;
                        case E2009.ST_PatternType.darkGrid:
                            return PatternType.DarkGrid;
                        case E2009.ST_PatternType.darkTrellis:
                            return PatternType.DarkTrellis;
                        case E2009.ST_PatternType.lightHorizontal:
                            return PatternType.LightHorizontal;
                        case E2009.ST_PatternType.lightVertical:
                            return PatternType.LightVertical;
                        case E2009.ST_PatternType.lightDown:
                            return PatternType.LightDown;
                        case E2009.ST_PatternType.lightUp:
                            return PatternType.LightUp;
                        case E2009.ST_PatternType.lightGrid:
                            return PatternType.LightGrid;
                        case E2009.ST_PatternType.lightTrellis:
                            return PatternType.LightTrellis;
                        case E2009.ST_PatternType.gray125:
                            return PatternType.Gray125;
                        case E2009.ST_PatternType.gray0625:
                            return PatternType.Gray0625;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                }
            }
            return null;
            //throw new ArgumentNullException(nameof(value));
        }

        private static Argb From2006Argb(string value)
        {
            if (value != null)
            {
                byte a, r, g, b;
                uint u;
                if (uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out u))
                {
                    b = (byte) (0xFF & (u >> 0));
                    g = (byte) (0xFF & (u >> 8));
                    r = (byte) (0xFF & (u >> 16));
                    a = (byte) (0xFF & (u >> 24));
                    return From2006Argb(a, r, g, b);
                }
                throw new InvalidEnumArgumentException(nameof(value));
            }
            return null;
        }

        private static Argb From2006Argb(byte a, byte r, byte g, byte b)
        {
            var result = new Argb
            {
                Value =
                {
                    [0] = a,
                    [1] = r,
                    [2] = g,
                    [3] = b
                }
            };
            return result;
        }
        #endregion

        #region 2009-RG

        public static ConditionalFormat From2009(E2009.CT_ConditionalFormatting value, IO.OpenXML.Document doc)
        {
            ConditionalFormat result = null;
            if (value != null)
            {
                result = new ConditionalFormat
                {
                    Pivot = value.pivot,
                    Sqref = From2009(value.sqref)
                };
                foreach (var rule in value.cfRule)
                {
                    result.Rules.Add(From2009(rule, doc));
                }
            }
            return result;
        }

        private static Sqref From2009(E2009.CT_Sqref value)
        {
            Sqref result = null;
            if (value != null)
            {
                result = new Sqref
                {
                    Adjusted = value.adjustedSpecified ? (bool?) value.adjusted : null,
                    Text = string.Concat(value.Text),
                    Edited = value.editedSpecified ? (bool?) value.edited : null,
                    Adjust = value.adjustSpecified ? (bool?) value.adjust : null,
                    Split = value.splitSpecified ? (bool?) value.split : null,
                };
            }
            return result;
        }

        private static ConditionalFormatRule From2009(E2009.CT_CfRule value, IO.OpenXML.Document doc)
        {
            ConditionalFormatRule result = null;
            if (value != null)
            {
                result = new ConditionalFormatRule
                {
                    AboveAverage = value.aboveAverage,
                    ActivePercent = value.activePresent,
                    Bottom = value.bottom,
                    ColorScale = From2009(value.colorScale, doc),
                    Text = value.text,
                    DataBar = From2009(value.dataBar, doc),
                    DifferentialFormat = From2009(value.dxf, doc),
                    EqualAverage = value.equalAverage,
                    IconSet = From2009(value.iconSet),
                    Operator = From2006(value.@operator, value.operatorSpecified),
                    Percent = value.percent,
                    Priority = value.prioritySpecified ? (int?) value.priority : null,
                    Rank = value.rankSpecified ? (uint?) value.rank : null,
                    SGuid = value.id,
                    StdDev = value.stdDevSpecified ? (int?) value.stdDev : null,
                    StopIfTrue = value.stopIfTrue,
                    TimePeriod = From2006(value.timePeriod, value.timePeriodSpecified),
                    Type = From2006(value.type, value.typeSpecified),
                };
                if (value.f != null)
                    foreach (var s in value.f)
                    {
                        result.Formula.Add(From2006(s));
                    }
            }
            return result;
        }

        private static ColorScale From2009(E2009.CT_ColorScale value, IO.OpenXML.Document doc)
        {
            ColorScale result = null;
            if (value != null)
            {
                result = new ColorScale
                {
                    
                };
                if (value.color != null)
                    foreach (var color in value.color)
                {
                    result.Color.Add(From2006(color, doc));
                }
                if (value.cfvo != null)
                    foreach (var cfvo in value.cfvo)
                {
                    result.CondittionalFormatValue.Add(From2009(cfvo));
                }
            }
            return result;
        }

        private static DataBar From2009(E2009.CT_DataBar value, IO.OpenXML.Document doc)
        {
            DataBar result = null;
            if (value != null)
            {
                result = new DataBar
                {
                    Border = value.border,
                    AxisColor = From2006(value.axisColor, doc),
                    AxisPosition = From2009(value.axisPosition),
                    BorderColor = From2006(value.borderColor, doc),
                    Direction = From2009(value.direction),
                    FillColor = From2006(value.fillColor, doc),
                    Gradient = value.gradient,
                    MaxLength = value.maxLength,
                    MinLength = value.minLength,
                    NegativeBarBorderColorSameAsPositive = value.negativeBarBorderColorSameAsPositive,
                    NegativeBarColorSameAsPositive = value.negativeBarColorSameAsPositive,
                    NegativeBorderColor = From2006(value.negativeBorderColor, doc),
                    NegativeFillColor = From2006(value.negativeFillColor, doc),
                    ShowValue = value.showValue,
                };
                if (value.cfvo != null)
                    foreach (var cfvo in value.cfvo)
                {
                    result.CondittionalFormatValue.Add(From2009(cfvo));
                }
            }
            return result;
        }

        private static DifferentialFormat From2009(E2009.CT_Dxf value, IO.OpenXML.Document doc)
        {
            DifferentialFormat result = null;
            if (value != null)
            {
                result = new DifferentialFormat
                {
                    Border = From2009(value.border, doc),
                    Font = From2009(value.font, doc),
                    NumberFormat = From2009(value.numFmt),
                    CellAlignment = From2009(value.alignment),
                    Fill = From2009(value.fill, doc),
                    CellProtection = From2009(value.protection),
                };
            }
            return result;
        }

        private static IconSet From2009(E2009.CT_IconSet value)
        {
            IconSet result = null;
            if (value != null)
            {
                result = new IconSet
                {
                    Percent = value.percent,
                    Custom = value.custom,
                    IconSetType = From2009(value.iconSet),
                    Reverse = value.reverse,
                    ShowValues = value.showValue,
                };
                if (value.cfvo != null)
                    foreach (var cfvo in value.cfvo)
                    {
                        result.CondittionalFormatValue.Add(From2009(cfvo));
                    }
                if (value.cfIcon != null)
                    foreach (var cfIcon in value.cfIcon)
                    {
                        result.Cficon.Add(From2009(cfIcon));
                    }
            }
            return result;
        }

        private static ConditionalFormatValueObject From2009(E2009.CT_Cfvo value)
        {
            ConditionalFormatValueObject result = null;
            if (value != null)
            {
                result = new ConditionalFormatValueObject
                {
                    Type = From2009(value.type),
                    Formula = From2006(value.f),
                    Gte = value.gte,
                };
            }
            return result;
        }

        //private static ConditionalFormatValueObject From2009(E2009.CT_Cfvo value)
        //{
        //    throw new NotImplementedException();
        //}

        //
        private static DatabarAxisPosition From2009(E2009.ST_DataBarAxisPosition value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_DataBarAxisPosition.automatic:
                        return DatabarAxisPosition.Automatic;
                    case E2009.ST_DataBarAxisPosition.middle:
                        return DatabarAxisPosition.Middle;
                    case E2009.ST_DataBarAxisPosition.none:
                        return DatabarAxisPosition.None;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }


        private static DataBarDirection From2009(E2009.ST_DataBarDirection value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_DataBarDirection.context:
                        return DataBarDirection.Context;
                    case E2009.ST_DataBarDirection.leftToRight:
                        return DataBarDirection.LeftToRight;
                    case E2009.ST_DataBarDirection.rightToLeft:
                        return DataBarDirection.RightToleft;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static Border From2009(E2009.CT_Border value, IO.OpenXML.Document doc)
        {
            Border result = new Border();
            if (value != null)
            {
                result = new Border
                {
                    Bottom = From2009(value.bottom, doc),
                    Outline = value.outline,
                    Vertical = From2009(value.vertical, doc),
                    Horizontal = From2009(value.horizontal, doc),
                    DiagonalDown = value.diagonalDownSpecified ? (bool?) value.diagonalDown : null,
                    Start = From2009(value.start, doc),
                    DiagonalUp = value.diagonalUpSpecified ? (bool?) value.diagonalUp : null,
                    Top = From2009(value.top, doc),
                    End = From2009(value.end, doc),
                    Diagonal = From2009(value.diagonal, doc),
                };
            }
            return result;
        }

        private static Font From2009(E2009.CT_Font value, IO.OpenXML.Document doc)
        {
            Font result = null;
            if (value != null)
            {
                result = new Font();
                var i = 0;
                foreach (var item in value.ItemsElementName)
                {
                    E2009.CT_BooleanProperty b;
                    switch (item)
                    {
                        case E2009.ItemsChoiceType.b:
                            b = value.Items[i] as E2009.CT_BooleanProperty;
                            if (b != null && b.val == false) b = null;
                            result.Bold = From2009(b);
                            break;
                        case E2009.ItemsChoiceType.charset:
                            result.Charset = From2009(value.Items[i] as E2009.CT_IntProperty);
                            break;
                        case E2009.ItemsChoiceType.color:
                            result.Color = From2006(value.Items[i] as E2006.CT_Color, doc);
                            break;
                        case E2009.ItemsChoiceType.condense:
                            b = value.Items[i] as E2009.CT_BooleanProperty;
                            if (b != null && b.val == false) b = null;
                            result.Condense = From2009(b);
                            break;
                        case E2009.ItemsChoiceType.extend:
                            b = value.Items[i] as E2009.CT_BooleanProperty;
                            if (b != null && b.val == false) b = null;
                            result.Extend = From2009(b);
                            break;
                        case E2009.ItemsChoiceType.family:
                            result.Family = From2009(value.Items[i] as E2009.CT_FontFamily);
                            break;
                        case E2009.ItemsChoiceType.i:
                            b = value.Items[i] as E2009.CT_BooleanProperty;
                            if (b != null && b.val == false) b = null;
                            result.Italic = From2009(b);
                            break;
                        case E2009.ItemsChoiceType.name:
                            result.Name = From2009(value.Items[i] as E2009.CT_FontName);
                            break;
                        case E2009.ItemsChoiceType.outline:
                            b = value.Items[i] as E2009.CT_BooleanProperty;
                            if (b != null && b.val == false) b = null;
                            result.Outline = From2009(b);
                            break;
                        case E2009.ItemsChoiceType.scheme:
                            result.FontScheme = From2009(value.Items[i] as E2009.CT_FontScheme);
                            break;
                        case E2009.ItemsChoiceType.shadow:
                            b = value.Items[i] as E2009.CT_BooleanProperty;
                            if (b != null && b.val == false) b = null;
                            result.Shadow = From2009(b);
                            break;
                        case E2009.ItemsChoiceType.strike:
                            result.Strike = From2009(value.Items[i] as E2009.CT_BooleanProperty);
                            break;
                        case E2009.ItemsChoiceType.sz:
                            result.FontSize = From2009(value.Items[i] as E2009.CT_FontSize);
                            break;
                        case E2009.ItemsChoiceType.u:
                            E2009.CT_UnderlineProperty u = value.Items[i] as E2009.CT_UnderlineProperty;
                            if (u != null && u.val == E2009.ST_UnderlineValues.none) u = null;
                            result.Underline = From2009(u);
                            break;
                        case E2009.ItemsChoiceType.vertAlign:
                            result.VerticalAlign = From2009(value.Items[i] as E2009.CT_VerticalAlignFontProperty);
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                    i++;
                }
            }
            return result;
        }

        private static NumberFormat From2009(E2009.CT_NumFmt value)
        {
            NumberFormat result = null;
            // так как мне не удалось понять как работает NumFmt - его импорт и экспорт игнорируются
            if (false && value != null)
            {
                result = new NumberFormat
                {
                    FormatCode = value.formatCode,
                    NumberFormatId = value.numFmtId,
                };
            }
            return result;
        }

        private static CellAlignment From2009(E2009.CT_CellAlignment value)
        {
            CellAlignment result = null;
            if (value != null)
            {
                result = new CellAlignment
                {
                    Horizontal = From2009(value.horizontal, value.horizontalSpecified),
                    Vertical = From2009(value.vertical),
                    JustifyLastLine = value.justifyLastLineSpecified ? (bool?) value.justifyLastLine : null,
                    Indent = value.indentSpecified ? (uint?) value.indent : null,
                    ReadingOrder = value.readingOrderSpecified ? (uint?) value.readingOrder : null,
                    RelativeIndent = value.relativeIndentSpecified ? (int?) value.relativeIndent : null,
                    ShrinkToFit = value.shrinkToFitSpecified ? (bool?) value.shrinkToFit : null,
                    TextRotation = value.textRotation,
                    WrapText = value.wrapTextSpecified ? (bool?) value.wrapText : null,
                };
            }
            return result;
        }

        private static Fill From2009(E2009.CT_Fill value, IO.OpenXML.Document doc)
        {
            Fill result = null;
            if (value != null)
            {
                result = new Fill
                {
                    GradientFill = From2009(value.Item as E2009.CT_GradientFill, doc),
                    PatternFill = From2009(value.Item as E2009.CT_PatternFill, doc),
                };
            }
            return result;
        }

        private static CellProtection From2009(E2009.CT_CellProtection value)
        {
            CellProtection result = null;
            if (value != null)
            {
                result = new CellProtection
                {
                    Locked = value.lockedSpecified?(bool?)value.locked:null,
                    Hidden = value.hiddenSpecified?(bool?)value.hidden:null,
                };
            }
            return result;
        }

        private static IconSetType From2009(E2009.ST_IconSetType value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_IconSetType.Item3Arrows:
                        return IconSetType.IconSet_3Arrows;
                    case E2009.ST_IconSetType.Item3ArrowsGray:
                        return IconSetType.IconSet_3ArrowsGray;
                    case E2009.ST_IconSetType.Item3Flags:
                        return IconSetType.IconSet_3Flags;
                    case E2009.ST_IconSetType.Item3TrafficLights1:
                        return IconSetType.IconSet_3TrafficLights1;
                    case E2009.ST_IconSetType.Item3TrafficLights2:
                        return IconSetType.IconSet_3TrafficLights2;
                    case E2009.ST_IconSetType.Item3Signs:
                        return IconSetType.IconSet_3Signs;
                    case E2009.ST_IconSetType.Item3Symbols:
                        return IconSetType.IconSet_3Symbols;
                    case E2009.ST_IconSetType.Item3Symbols2:
                        return IconSetType.IconSet_3Symbols2;
                    case E2009.ST_IconSetType.Item4Arrows:
                        return IconSetType.IconSet_4Arrows;
                    case E2009.ST_IconSetType.Item4ArrowsGray:
                        return IconSetType.IconSet_4ArrowsGray;
                    case E2009.ST_IconSetType.Item4RedToBlack:
                        return IconSetType.IconSet_4RedToBlack;
                    case E2009.ST_IconSetType.Item4Rating:
                        return IconSetType.IconSet_4Rating;
                    case E2009.ST_IconSetType.Item4TrafficLights:
                        return IconSetType.IconSet_4TrafficLights;
                    case E2009.ST_IconSetType.Item5Arrows:
                        return IconSetType.IconSet_5Arrows;
                    case E2009.ST_IconSetType.Item5ArrowsGray:
                        return IconSetType.IconSet_5ArrowsGray;
                    case E2009.ST_IconSetType.Item5Rating:
                        return IconSetType.IconSet_5Rating;
                    case E2009.ST_IconSetType.Item5Quarters:
                        return IconSetType.IconSet_5Quarters;
                    case E2009.ST_IconSetType.Item3Stars:
                        return  IconSetType.IconSet_3Stars;
                    case E2009.ST_IconSetType.Item3Triangles:
                        return IconSetType.IconSet_3Triangles;
                    case E2009.ST_IconSetType.Item5Boxes:
                        return IconSetType.IconSet_5Boxes;
                    case E2009.ST_IconSetType.NoIcons:
                        return IconSetType.IconSet_NoIcons;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static CfIcon From2009(E2009.CT_CfIcon value)
        {
            CfIcon result = null;
            if (value != null)
            {
                result = new CfIcon
                {
                    IconId = value.iconId,
                    IconSet = From2009(value.iconSet),
                };
            }
            return result;
        }


        private static ConditionalFormatValueObjectType From2009(E2009.ST_CfvoType value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_CfvoType.num:
                        return ConditionalFormatValueObjectType.Num;
                    case E2009.ST_CfvoType.percent:
                        return ConditionalFormatValueObjectType.Percent;
                    case E2009.ST_CfvoType.max:
                        return ConditionalFormatValueObjectType.Max;
                    case E2009.ST_CfvoType.min:
                        return ConditionalFormatValueObjectType.Min;
                    case E2009.ST_CfvoType.formula:
                        return ConditionalFormatValueObjectType.Formula;
                    case E2009.ST_CfvoType.percentile:
                        return ConditionalFormatValueObjectType.Percentile;
                    case E2009.ST_CfvoType.autoMin:
                        return ConditionalFormatValueObjectType.AutoMin;
                    case E2009.ST_CfvoType.autoMax:
                        return ConditionalFormatValueObjectType.AutoMax;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static BorderPr From2009(E2009.CT_BorderPr value, IO.OpenXML.Document doc)
        {
            BorderPr result = null;
            if (value != null)
            {
                result = new BorderPr
                {
                    Color = From2006(value.color, doc),
                    Style = From2009(value.style),
                };
            }
            return result;
        }

        private static BooleanProperty From2009(E2009.CT_BooleanProperty value)
        {
            BooleanProperty result = null;
            if (value != null)
            {
                result = new BooleanProperty
                {
                    Value = value.val
                };
            }
            return result;
        }

        private static IntProperty From2009(E2009.CT_IntProperty value)
        {
            IntProperty result = null;
            if (value != null)
            {
                result = new IntProperty
                {
                    Value = value.val
                };
            }
            return result;
        }

        private static FontFamily From2009(E2009.CT_FontFamily value)
        {
            FontFamily result = null;
            if (value != null)
            {
                result = new FontFamily
                {
                    Value = value.val
                };
            }
            return result;
        }

        private static FontName From2009(E2009.CT_FontName value)
        {
            FontName result = null;
            if (value != null)
            {
                result = new FontName
                {
                    Value = value.val,
                };
            }
            return result;
        }

        private static FontScheme From2009(E2009.CT_FontScheme value)
        {
            FontScheme result = null;
            if (value != null)
            {
                result = new FontScheme
                {
                    Value = From2009(value.val),
                };
            }
            return result;
        }

        private static FontSize From2009(E2009.CT_FontSize value)
        {
            FontSize result = null;
            if (value != null)
            {
                result = new FontSize
                {
                    Value = value.val
                };
            }
            return result;
        }

        private static UnderlineProperty From2009(E2009.CT_UnderlineProperty value)
        {
            UnderlineProperty result = null;
            if (value != null)
            {
                result = new UnderlineProperty
                {
                    Value = From2009(value.val),
                };
            }
            return result;
        }

        private static VerticalAlignFontProperty From2009(E2009.CT_VerticalAlignFontProperty value)
        {
            VerticalAlignFontProperty result = null;
            if (value != null)
            {
                result = new VerticalAlignFontProperty
                {
                    Value = From2009(value.val),
                };
            }
            return result;
        }

        private static HorizontalAlignment? From2009(E2009.ST_HorizontalAlignment value, bool specified)
        {
            if (!specified) return null;
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_HorizontalAlignment.general:
                        return HorizontalAlignment.General;
                    case E2009.ST_HorizontalAlignment.left:
                        return HorizontalAlignment.Left;
                    case E2009.ST_HorizontalAlignment.center:
                        return HorizontalAlignment.Center;
                    case E2009.ST_HorizontalAlignment.right:
                        return HorizontalAlignment.Right;
                    case E2009.ST_HorizontalAlignment.fill:
                        return HorizontalAlignment.Fill;
                    case E2009.ST_HorizontalAlignment.justify:
                        return HorizontalAlignment.Justify;
                    case E2009.ST_HorizontalAlignment.centerContinuous:
                        return HorizontalAlignment.CenterContinuous;
                    case E2009.ST_HorizontalAlignment.distributed:
                        return HorizontalAlignment.Distributed;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static VerticalAlignment From2009(E2009.ST_VerticalAlignment value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_VerticalAlignment.top:
                        return VerticalAlignment.Top;
                    case E2009.ST_VerticalAlignment.center:
                        return VerticalAlignment.Center;
                    case E2009.ST_VerticalAlignment.bottom:
                        return VerticalAlignment.Bottom;
                    case E2009.ST_VerticalAlignment.justify:
                        return VerticalAlignment.Justify;
                    case E2009.ST_VerticalAlignment.distributed:
                        return VerticalAlignment.Distributed;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static GradientFill From2009(E2009.CT_GradientFill value, IO.OpenXML.Document doc)
        {
            GradientFill result = null;
            if (value != null)
            {
                result = new GradientFill
                {
                    Bottom = value.bottom,
                    Top = value.top,
                    Degree = value.degree,
                    GradientType = From2009(value.type),
                    Left = value.left,
                    Right = value.right,
                };
                if (value.stop != null)
                    foreach (var stop in value.stop)
                {
                    result.GradientStop.Add(From2009(stop, doc));
                }
            }
            return result;
        }

        private static PatternFill From2009(E2009.CT_PatternFill value, IO.OpenXML.Document doc)
        {
            PatternFill result = null;
            if (value != null)
            {
                result = new PatternFill
                {
                    PatternType = From2009(value.patternType, value.patternTypeSpecified),
                    ForegroundColor = From2006(value.fgColor, doc),
                    BackgroundColor = From2006(value.bgColor, doc),
                };
            }
            return result;
        }

        private static BorderStyle From2009(E2009.ST_BorderStyle value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_BorderStyle.none:
                        return BorderStyle.None;
                    case E2009.ST_BorderStyle.thin:
                        return BorderStyle.Thin;
                    case E2009.ST_BorderStyle.medium:
                        return BorderStyle.Medium;
                    case E2009.ST_BorderStyle.dashed:
                        return BorderStyle.Dashed;
                    case E2009.ST_BorderStyle.dotted:
                        return BorderStyle.Dotted;
                    case E2009.ST_BorderStyle.thick:
                        return BorderStyle.Thick;
                    case E2009.ST_BorderStyle.@double:
                        return BorderStyle.Double;
                    case E2009.ST_BorderStyle.hair:
                        return BorderStyle.Hair;
                    case E2009.ST_BorderStyle.mediumDashed:
                        return BorderStyle.MediumDashed;
                    case E2009.ST_BorderStyle.dashDot:
                        return BorderStyle.DashDot;
                    case E2009.ST_BorderStyle.mediumDashDot:
                        return BorderStyle.MediumDashDot;
                    case E2009.ST_BorderStyle.dashDotDot:
                        return BorderStyle.DashDotDot;
                    case E2009.ST_BorderStyle.mediumDashDotDot:
                        return BorderStyle.MediumDashDotDot;
                    case E2009.ST_BorderStyle.slantDashDot:
                        return BorderStyle.SlantDashDot;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static FontSchemeEnum From2009(E2009.ST_FontScheme value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_FontScheme.none:
                        return FontSchemeEnum.None;
                    case E2009.ST_FontScheme.major:
                        return FontSchemeEnum.Major;
                    case E2009.ST_FontScheme.minor:
                        return FontSchemeEnum.Minor;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static UnderlineValues From2009(E2009.ST_UnderlineValues value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_UnderlineValues.single:
                        return UnderlineValues.Single;
                    case E2009.ST_UnderlineValues.@double:
                        return UnderlineValues.Double;
                    case E2009.ST_UnderlineValues.singleAccounting:
                        return UnderlineValues.SingleAccounting;
                    case E2009.ST_UnderlineValues.doubleAccounting:
                        return UnderlineValues.DoubleAccounting;
                    case E2009.ST_UnderlineValues.none:
                        return UnderlineValues.None;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static VerticalAlignRun From2009(E2009.ST_VerticalAlignRun value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_VerticalAlignRun.baseline:
                        return  VerticalAlignRun.Baseline;
                    case E2009.ST_VerticalAlignRun.superscript:
                        return VerticalAlignRun.Superscript;
                    case E2009.ST_VerticalAlignRun.subscript:
                        return VerticalAlignRun.Subscript;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static GradientType From2009(E2009.ST_GradientType value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_GradientType.linear:
                        return GradientType.Linear;
                    case E2009.ST_GradientType.path:
                        return GradientType.Path;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        private static GradientStop From2009(E2009.CT_GradientStop value, IO.OpenXML.Document doc)
        {
            GradientStop result = null;
            if (value != null)
            {
                result = new GradientStop
                {
                    Color = From2006(value.color, doc),
                    Position = value.position,
                };
            }
            return result;
        }

        private static PatternType? From2009(E2009.ST_PatternType value, bool specified)
        {
            if (!specified) return null;
            if (value != null)
            {
                switch (value)
                {
                    case E2009.ST_PatternType.none:
                        return PatternType.None;
                    case E2009.ST_PatternType.solid:
                        return PatternType.Solid;
                    case E2009.ST_PatternType.mediumGray:
                        return PatternType.MediumGray;
                    case E2009.ST_PatternType.darkGray:
                        return PatternType.DarkGray;
                    case E2009.ST_PatternType.lightGray:
                        return PatternType.LightGray;
                    case E2009.ST_PatternType.darkHorizontal:
                        return PatternType.DarkHorizontal;
                    case E2009.ST_PatternType.darkVertical:
                        return PatternType.DarkVertical;
                    case E2009.ST_PatternType.darkDown:
                        return PatternType.DarkDown;
                    case E2009.ST_PatternType.darkUp:
                        return PatternType.DarkUp;
                    case E2009.ST_PatternType.darkGrid:
                        return PatternType.DarkGrid;
                    case E2009.ST_PatternType.darkTrellis:
                        return PatternType.DarkTrellis;
                    case E2009.ST_PatternType.lightHorizontal:
                        return PatternType.LightHorizontal;
                    case E2009.ST_PatternType.lightVertical:
                        return PatternType.LightVertical;
                    case E2009.ST_PatternType.lightDown:
                        return PatternType.LightDown;
                    case E2009.ST_PatternType.lightUp:
                        return PatternType.LightUp;
                    case E2009.ST_PatternType.lightGrid:
                        return PatternType.LightGrid;
                    case E2009.ST_PatternType.lightTrellis:
                        return PatternType.LightTrellis;
                    case E2009.ST_PatternType.gray125:
                        return PatternType.Gray125;
                    case E2009.ST_PatternType.gray0625:
                        return PatternType.Gray0625;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, null);
                }
            }
            throw new ArgumentNullException(nameof(value));
        }

        #endregion

        #region RG->2009

        private static E2009.CT_ConditionalFormatting ToExcel2009(ConditionalFormat value)
        {

            E2009.CT_ConditionalFormatting result = null;
            if (value != null)
            {
                result = new E2009.CT_ConditionalFormatting
                {
                    pivot = value.Pivot,
                    sqref = ToExcel2009(value.Sqref),
                    extLst = null, // WARNING: расширения игнорируются!
                };
                if (value.Rules != null && value.Rules.Count > 0)
                {
                    var list = new List<E2009.CT_CfRule>();
                    foreach (var rule in value.Rules)
                    {
                        var item = ToExcel2009(rule);
                        if(item != null)
                            list.Add(item);
                    }
                    if (list.Count > 0)
                        result.cfRule = list.ToArray();
                }
            }
            return result;
        }

        private static E2009.CT_Sqref ToExcel2009(Sqref value)
        {
            E2009.CT_Sqref result = null;
            if (value != null)
            {
                result = new E2009.CT_Sqref
                {
                    adjust = value.Adjust ?? false,
                    adjustSpecified = value.Adjust != null,
                    adjusted = value.Adjusted ?? false,
                    adjustedSpecified = value.Adjusted != null,
                    edited = value.Edited ?? false,
                    editedSpecified = value.Edited != null,
                    split = value.Split ?? false,
                    splitSpecified = value.Edited != null,
                    Text = new[] {value.Text},
                };
            }
            return result;
        }

        private static E2009.CT_CfRule ToExcel2009(ConditionalFormatRule value)
        {
            E2009.CT_CfRule result = null;
            if (value != null)
            {
                bool operatorSpecified;
                var @operator = ToExcel2009(value.Operator, out operatorSpecified);
                bool typSpecified;
                var type = ToExcel2009(value.Type, out typSpecified);
                bool timePeriodSpecified;
                var timePeriod = ToExcel2009(value.TimePeriod, out timePeriodSpecified);

                result = new E2009.CT_CfRule
                {
                    @operator = @operator,
                    operatorSpecified = operatorSpecified,
                    dxf = ToExcel2009(value.DifferentialFormat),
                    extLst = null,                          // WARNING расширения теряются
                    stopIfTrue = value.StopIfTrue ?? false,
                    stdDev = value.StdDev??0,
                    stdDevSpecified = value.StdDev != null,
                    iconSet = ToExcel2009(value.IconSet),
                    aboveAverage = value.AboveAverage??true,
                    equalAverage = value.EqualAverage??false,
                    percent = value.Percent??false,
                    dataBar = ToExcel2009(value.DataBar),
                    bottom = value.Bottom ?? false,
                    type = type,
                    typeSpecified = typSpecified,
                    rank = value.Rank??0,
                    rankSpecified =  value.Rank != null,
                    priority = value.Priority ?? 0,         //TODO тут у меня разные xsd и документация не сходится вроде он должен быть required
                    prioritySpecified = value.Priority != null,
                    timePeriod = timePeriod,
                    timePeriodSpecified = timePeriodSpecified,
                    activePresent = value.ActivePercent??false,
                    colorScale = ToExcel2009(value.ColorScale),
                    id = value.SGuid,
                    text = value.Text,
                };
                var list = new List<string>();
                foreach (var item in value.Formula)
                {
                    var f = item.Value;
                    if(f != null)
                        list.Add(f);
                }
                result.f = list.ToArray();

            }
            return result;
        }

        private static E2006.ST_ConditionalFormattingOperator ToExcel2009(ConditionalFormattingOperator? value,
            out bool specified)
        {
            specified = value != null;
            switch (value)
            {
                case ConditionalFormattingOperator.LessThan:
                    return E2006.ST_ConditionalFormattingOperator.lessThan;
                case ConditionalFormattingOperator.LessThanOrEqual:
                    return E2006.ST_ConditionalFormattingOperator.lessThanOrEqual;
                case ConditionalFormattingOperator.Equal:
                    return E2006.ST_ConditionalFormattingOperator.equal;
                case ConditionalFormattingOperator.NotEqual:
                    return E2006.ST_ConditionalFormattingOperator.notEqual;
                case ConditionalFormattingOperator.GreaterThanOrEqual:
                    return E2006.ST_ConditionalFormattingOperator.greaterThanOrEqual;
                case ConditionalFormattingOperator.GreaterThan:
                    return E2006.ST_ConditionalFormattingOperator.greaterThan;
                case ConditionalFormattingOperator.Between:
                    return E2006.ST_ConditionalFormattingOperator.between;
                case ConditionalFormattingOperator.NotBetween:
                    return E2006.ST_ConditionalFormattingOperator.notBetween;
                case ConditionalFormattingOperator.ContainsText:
                    return E2006.ST_ConditionalFormattingOperator.containsText;
                case ConditionalFormattingOperator.NotContains:
                    return E2006.ST_ConditionalFormattingOperator.notContains;
                case ConditionalFormattingOperator.BeginsWith:
                    return E2006.ST_ConditionalFormattingOperator.beginsWith;
                case ConditionalFormattingOperator.EndsWith:
                    return E2006.ST_ConditionalFormattingOperator.endsWith;
                case null:
                    return E2006.ST_ConditionalFormattingOperator.lessThan;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.CT_Dxf ToExcel2009(DifferentialFormat value)
        {
            E2009.CT_Dxf result = null;
            if (value != null)
            {
                result = new E2009.CT_Dxf
                {
                    extLst = null,                          // WARNING расширения теряются
                    font = ToExcel2009(value.Font),
                    border = ToExcel2009(value.Border),
                    fill = ToExcel2009(value.Fill),
                    protection = ToExcel2009(value.CellProtection),
                    alignment = ToExcel2009(value.CellAlignment),
                    numFmt = ToExcel2009(value.NumberFormat),
                };
            }
            return result;
        }

        private static E2009.CT_IconSet ToExcel2009(IconSet value)
        {
            E2009.CT_IconSet result = null;
            if (value != null)
            {
                result = new E2009.CT_IconSet
                {
                    custom = value.Custom ?? false,
                    iconSet = ToExcel2009(value.IconSetType),
                    percent = value.Percent ?? true,
                    reverse = value.Reverse ?? false,
                    showValue = value.ShowValues ?? true,
                    
                };
                {
                    var list = new List<E2009.CT_Cfvo>();
                    if(value.CondittionalFormatValue != null )
                    foreach (var cfvalue in value.CondittionalFormatValue)
                    {
                        var item = ToExcel2009(cfvalue);
                        if (item != null)
                        {
                            list.Add(item);
                        }
                    }
                    result.cfvo = list.ToArray();
                }
                {
                    var list = new List<E2009.CT_CfIcon>();
                    if (value.Cficon != null)
                        foreach (var cfvalue in value.Cficon)
                        {
                            var item = ToExcel2009(cfvalue);
                            if (item != null)
                            {
                                list.Add(item);
                            }
                        }
                    result.cfIcon = list.ToArray();
                }
            }
            return result;
        }

        private static E2009.CT_DataBar ToExcel2009(DataBar value)
        {
            E2009.CT_DataBar result = null;
            if (value != null)
            {
                result = new E2009.CT_DataBar
                {
                    axisColor = ToExcel2009(value.AxisColor),
                    axisPosition = ToExcel2009(value.AxisPosition ?? DatabarAxisPosition.Automatic),
                    border = value.Border ?? false,
                    borderColor = ToExcel2009(value.BorderColor),
                    direction = ToExcel2009(value.Direction ?? DataBarDirection.Context),
                    fillColor = ToExcel2009(value.FillColor),
                    gradient = value.Gradient ?? true,
                    maxLength = value.MaxLength ?? 90,
                    minLength = value.MinLength ?? 10,
                    negativeBarBorderColorSameAsPositive = value.NegativeBarBorderColorSameAsPositive ?? true,
                    negativeBarColorSameAsPositive = value.NegativeBarColorSameAsPositive ?? false,
                    negativeBorderColor = ToExcel2009(value.NegativeBorderColor),
                    negativeFillColor = ToExcel2009(value.NegativeFillColor),
                    showValue = value.ShowValue ?? true,
                };
                {
                    var list = new List<E2009.CT_Cfvo>();
                    if (value.CondittionalFormatValue != null)
                        foreach (var cfvalue in value.CondittionalFormatValue)
                        {
                            var item = ToExcel2009(cfvalue);
                            if (item != null)
                            {
                                list.Add(item);
                            }
                        }
                    result.cfvo = list.ToArray();
                }
            }
            return result;
        }

        private static E2006.ST_CfType ToExcel2009(ConditionalFormatType? value, out bool specified)
        {
            specified = value != null;
            switch (value)
            {
                case ConditionalFormatType.Expression:
                    return E2006.ST_CfType.expression;
                case ConditionalFormatType.CellIs:
                    return E2006.ST_CfType.cellIs;
                case ConditionalFormatType.ColorScale:
                    return E2006.ST_CfType.colorScale;
                case ConditionalFormatType.DataBar:
                    return E2006.ST_CfType.dataBar;
                case ConditionalFormatType.IconSet:
                    return E2006.ST_CfType.iconSet;
                case ConditionalFormatType.Top10:
                    return E2006.ST_CfType.top10;
                case ConditionalFormatType.UniqueValues:
                    return E2006.ST_CfType.uniqueValues;
                case ConditionalFormatType.DuplicateValues:
                    return E2006.ST_CfType.duplicateValues;
                case ConditionalFormatType.ContainsText:
                    return E2006.ST_CfType.containsText;
                case ConditionalFormatType.NotContainsText:
                    return E2006.ST_CfType.notContainsText;
                case ConditionalFormatType.BeginsWith:
                    return E2006.ST_CfType.beginsWith;
                case ConditionalFormatType.EndsWith:
                    return E2006.ST_CfType.endsWith;
                case ConditionalFormatType.ContainsBlanks:
                    return E2006.ST_CfType.containsBlanks;
                case ConditionalFormatType.NotContainsBlanks:
                    return E2006.ST_CfType.notContainsBlanks;
                case ConditionalFormatType.ContainsErrors:
                    return E2006.ST_CfType.containsErrors;
                case ConditionalFormatType.NotContainsErrors:
                    return E2006.ST_CfType.notContainsErrors;
                case ConditionalFormatType.TimePeriod:
                    return E2006.ST_CfType.timePeriod;
                case ConditionalFormatType.AboveAverage:
                    return E2006.ST_CfType.aboveAverage;
                case null:
                    return E2006.ST_CfType.expression;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2006.ST_TimePeriod ToExcel2009(TimePeriod? value, out bool specified)
        {
            specified = value != null;
            switch (value)
            {
                case TimePeriod.Today:
                    return E2006.ST_TimePeriod.today;
                case TimePeriod.Yesterday:
                    return E2006.ST_TimePeriod.yesterday;
                case TimePeriod.Tomorrow:
                    return E2006.ST_TimePeriod.tomorrow;
                case TimePeriod.Last7Days:
                    return E2006.ST_TimePeriod.last7Days;
                case TimePeriod.ThisMonth:
                    return E2006.ST_TimePeriod.thisMonth;
                case TimePeriod.LastMonth:
                    return E2006.ST_TimePeriod.lastMonth;
                case TimePeriod.NextMonth:
                    return E2006.ST_TimePeriod.nextMonth;
                case TimePeriod.ThisWeek:
                    return E2006.ST_TimePeriod.thisWeek;
                case TimePeriod.LastWeek:
                    return E2006.ST_TimePeriod.lastWeek;
                case TimePeriod.NextWeek:
                    return E2006.ST_TimePeriod.nextWeek;
                case null:
                    return E2006.ST_TimePeriod.today;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.CT_ColorScale ToExcel2009(ColorScale value)
        {
            E2009.CT_ColorScale result = null;
            if (value != null)
            {
                result = new E2009.CT_ColorScale
                {
                };
                {
                    var list = new List<E2009.CT_Cfvo>();
                    if (value.CondittionalFormatValue != null)
                        foreach (var cfvalue in value.CondittionalFormatValue)
                        {
                            var item = ToExcel2009(cfvalue);
                            if (item != null)
                            {
                                list.Add(item);
                            }
                        }
                    result.cfvo = list.ToArray();
                }
                {
                    var list = new List<E2006.CT_Color>();
                    if (value.CondittionalFormatValue != null)
                        foreach (var cfvalue in value.Color)
                        {
                            var item = ToExcel2009(cfvalue);
                            if (item != null)
                            {
                                list.Add(item);
                            }
                        }
                    result.color = list.ToArray();
                }
            }
            return result;
        }

        private static E2009.CT_Font ToExcel2009(Font value)
        {
            E2009.CT_Font result = null;
            if (value != null)
            {
                result = new E2009.CT_Font
                {
                };
                {
                    var list = new List<object>();
                    var names = new List<E2009.ItemsChoiceType>();
                    var add = new Action<object, E2009.ItemsChoiceType>((o, t) =>
                    {
                        if (o != null)
                        {
                            list.Add(o);
                            names.Add(t);
                        }
                    });
                    if (value?.Bold?.Value == true)
                        add(ToExcel2009(value.Bold), E2009.ItemsChoiceType.b);
                    add(ToExcel2009(value.Charset), E2009.ItemsChoiceType.charset);
                    add(ToExcel2009(value.Color), E2009.ItemsChoiceType.color);
                    if (value?.Condense?.Value == true)
                        add(ToExcel2009(value.Condense), E2009.ItemsChoiceType.condense);
                    if (value?.Extend?.Value == true)
                        add(ToExcel2009(value.Extend), E2009.ItemsChoiceType.extend);
                    add(ToExcel2009(value.Family), E2009.ItemsChoiceType.family);
                    if (value?.Italic?.Value == true)
                        add(ToExcel2009(value.Italic), E2009.ItemsChoiceType.i);
                    add(ToExcel2009(value.Name), E2009.ItemsChoiceType.name);
                    if (value?.Outline?.Value == true)
                        add(ToExcel2009(value.Outline), E2009.ItemsChoiceType.outline);
                    add(ToExcel2009(value.FontScheme), E2009.ItemsChoiceType.scheme);
                    if (value?.Shadow?.Value == true)
                        add(ToExcel2009(value.Shadow), E2009.ItemsChoiceType.shadow);
                    if (value?.Strike?.Value == true)
                        add(ToExcel2009(value.Strike), E2009.ItemsChoiceType.strike);
                    add(ToExcel2009(value.FontSize), E2009.ItemsChoiceType.sz);
                    if (value?.Underline?.Value != null &&
                        value?.Underline?.Value != UnderlineValues.None)
                        add(ToExcel2009(value.Underline), E2009.ItemsChoiceType.u);
                    add(ToExcel2009(value.VerticalAlign), E2009.ItemsChoiceType.vertAlign);

                    result.Items = list.ToArray();
                    result.ItemsElementName = names.ToArray();
                }
            }
            return result;
        }

        private static E2009.CT_Border ToExcel2009(Border value)
        {
            E2009.CT_Border result = null;
            if (value != null)
            {
                result = new E2009.CT_Border
                {
                    horizontal = ToExcel2009(value.Horizontal),
                    bottom = ToExcel2009(value.Bottom),
                    diagonal = ToExcel2009(value.Diagonal),
                    end = ToExcel2009(value.End),
                    start = ToExcel2009(value.Start),
                    top = ToExcel2009(value.Top),
                    vertical = ToExcel2009(value.Vertical),
                    diagonalDown = value.DiagonalDown ?? false,
                    diagonalDownSpecified = value.DiagonalDown != null,
                    diagonalUp = value.DiagonalUp ?? false,
                    diagonalUpSpecified = value.DiagonalUp != null,
                    outline = value.Outline ?? true,
                };
            }
            return result;
        }

        private static E2009.CT_Fill ToExcel2009(Fill value)
        {
            E2009.CT_Fill result = null;
            if (CanExportExcel(value) )
            {
                result = new E2009.CT_Fill
                {
                    Item = (object)ToExcel2009(value.PatternFill) ?? ToExcel2009(value.GradientFill),
                };
            }
            return result;
        }

        private static bool CanExportExcel(Fill value)
        {
            if (value == null) return false;
            if (value.PatternFill?.BackgroundColor?.RgbColorValue != null) return true;
            if (value.GradientFill != null) return true;
            return false;
        }

        private static E2009.CT_CellProtection ToExcel2009(CellProtection value)
        {
            E2009.CT_CellProtection result = null;
            if (value != null)
            {
                result = new E2009.CT_CellProtection
                {
                    locked = value.Locked??false,
                    hidden = value.Hidden??false,
                    hiddenSpecified = value.Hidden != null,
                    lockedSpecified = value.Locked != null,
                };
            }
            return result;
        }

        private static E2009.CT_CellAlignment ToExcel2009(CellAlignment value)
        {
            E2009.CT_CellAlignment result = null;
            if (value != null)
            {
                bool horizontalSpecified;
                var horizontal = ToExcel2009(value.Horizontal, out horizontalSpecified);

                result = new E2009.CT_CellAlignment
                {
                    indent = value.Indent??0,
                    indentSpecified = value.Indent != null,
                    wrapText = value.WrapText??false,
                    wrapTextSpecified = value.WrapText != null,
                    readingOrder = value.ReadingOrder??0,
                    readingOrderSpecified = value.ReadingOrder != null,
                    justifyLastLine = value.JustifyLastLine??false,
                    justifyLastLineSpecified = value.JustifyLastLine != null,
                    relativeIndent = value.RelativeIndent??0,
                    relativeIndentSpecified = value.RelativeIndent != null,
                    shrinkToFit = value.ShrinkToFit??false,
                    shrinkToFitSpecified = value.ShrinkToFit != null,
                    horizontal = horizontal,
                    horizontalSpecified = horizontalSpecified,
                    vertical = ToExcel2009(value.Vertical),
                    textRotation = value.TextRotation,
                };
            }
            return result;
        }

        private static E2009.CT_NumFmt ToExcel2009(NumberFormat value)
        {
            E2009.CT_NumFmt result = null;
            // так как мне не удалось понять как работает NumFmt - его импорт и экспорт игнорируются
            if (false && value != null)
            {
                result = new E2009.CT_NumFmt
                {
                    numFmtId = value.NumberFormatId,        // TODO проврить что за формат
                    formatCode = value.FormatCode,
                };
            }
            return result;
        }

        private static E2009.ST_IconSetType ToExcel2009(IconSetType value)
        {
            switch (value)
            {
                case IconSetType.IconSet_3Arrows:
                    return E2009.ST_IconSetType.Item3Arrows;
                case IconSetType.IconSet_3ArrowsGray:
                    return E2009.ST_IconSetType.Item3ArrowsGray;
                case IconSetType.IconSet_3Flags:
                    return E2009.ST_IconSetType.Item3Flags;
                case IconSetType.IconSet_3TrafficLights1:
                    return E2009.ST_IconSetType.Item3TrafficLights1;
                case IconSetType.IconSet_3TrafficLights2:
                    return E2009.ST_IconSetType.Item3TrafficLights2;
                case IconSetType.IconSet_3Signs:
                    return E2009.ST_IconSetType.Item3Signs;
                case IconSetType.IconSet_3Symbols:
                    return E2009.ST_IconSetType.Item3Symbols;
                case IconSetType.IconSet_3Symbols2:
                    return E2009.ST_IconSetType.Item3Symbols2;
                case IconSetType.IconSet_4Arrows:
                    return E2009.ST_IconSetType.Item4Arrows;
                case IconSetType.IconSet_4ArrowsGray:
                    return E2009.ST_IconSetType.Item4ArrowsGray;
                case IconSetType.IconSet_4RedToBlack:
                    return E2009.ST_IconSetType.Item4RedToBlack;
                case IconSetType.IconSet_4Rating:
                    return E2009.ST_IconSetType.Item4Rating;
                case IconSetType.IconSet_4TrafficLights:
                    return E2009.ST_IconSetType.Item4TrafficLights;
                case IconSetType.IconSet_5Arrows:
                    return E2009.ST_IconSetType.Item5Arrows;
                case IconSetType.IconSet_5ArrowsGray:
                    return E2009.ST_IconSetType.Item5ArrowsGray;
                case IconSetType.IconSet_5Rating:
                    return E2009.ST_IconSetType.Item5Rating;
                case IconSetType.IconSet_5Quarters:
                    return E2009.ST_IconSetType.Item5Quarters;
                case IconSetType.IconSet_3Stars:
                    return E2009.ST_IconSetType.Item3Stars;
                case IconSetType.IconSet_3Triangles:
                    return E2009.ST_IconSetType.Item3Triangles;
                case IconSetType.IconSet_5Boxes:
                    return E2009.ST_IconSetType.Item5Boxes;
                case IconSetType.IconSet_NoIcons:
                    return E2009.ST_IconSetType.NoIcons;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.CT_Cfvo ToExcel2009(ConditionalFormatValueObject value)
        {
            E2009.CT_Cfvo result = null;
            if (value != null)
            {
                result = new E2009.CT_Cfvo
                {
                    extLst = null,
                    type = ToExcel2009(value.Type),
                    f = value.Formula?.Value,
                    gte = value.Gte,
                };
            }
            return result;
        }

        private static E2009.CT_CfIcon ToExcel2009(CfIcon value)
        {
            E2009.CT_CfIcon result = null;
            if (value != null)
            {
                result = new E2009.CT_CfIcon
                {
                    iconSet = ToExcel2009(value.IconSet),
                    iconId = value.IconId,
                };
            }
            return result;
        }

        private static E2006.CT_Color ToExcel2009(Color value)
        {
            E2006.CT_Color result = null;
            if (value != null)
            {
                result = new E2006.CT_Color
                {
                    tint = value.RgbColorValue != null ? 0D : value.TInt ?? 0D,
                    rgb = ToExcel2009ArgbToRgb(value.RgbColorValue),
                    indexed = value.Indexed??0,
                    indexedSpecified = value.Indexed != null,
                    auto = value.Automatic??false,
                    autoSpecified = value.Automatic != null,
                    theme =  value.ThemeColor??0,
                    themeSpecified = value.ThemeColor != null,
                };
            }
            return result;
        }

        private static E2009.ST_DataBarAxisPosition ToExcel2009(DatabarAxisPosition value)
        {
            switch (value)
            {
                case DatabarAxisPosition.Automatic:
                    return E2009.ST_DataBarAxisPosition.automatic;
                case DatabarAxisPosition.Middle:
                    return E2009.ST_DataBarAxisPosition.middle;
                case DatabarAxisPosition.None:
                    return E2009.ST_DataBarAxisPosition.none;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.ST_DataBarDirection ToExcel2009(DataBarDirection value)
        {
            switch (value)
            {
                case DataBarDirection.Context:
                    return E2009.ST_DataBarDirection.context;
                case DataBarDirection.LeftToRight:
                    return E2009.ST_DataBarDirection.leftToRight;
                case DataBarDirection.RightToleft:
                    return E2009.ST_DataBarDirection.rightToLeft;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.CT_BooleanProperty ToExcel2009(BooleanProperty value)
        {
            E2009.CT_BooleanProperty result = null;
            if (value != null)
            {
                result = new E2009.CT_BooleanProperty
                {
                    val = value.Value
                };
            }
            return result;
        }

        private static E2009.CT_IntProperty ToExcel2009(IntProperty value)
        {
            E2009.CT_IntProperty result = null;
            if (value != null)
            {
                result = new E2009.CT_IntProperty
                {
                    val = value.Value,
                };
            }
            return result;
        }

        private static E2009.CT_FontFamily ToExcel2009(FontFamily value)
        {
            E2009.CT_FontFamily result = null;
            if (value != null)
            {
                result = new E2009.CT_FontFamily
                {
                    val = value.Value,
                };
            }
            return result;
        }

        private static E2009.CT_FontName ToExcel2009(FontName value)
        {
            E2009.CT_FontName result = null;
            if (value != null)
            {
                result = new E2009.CT_FontName
                {
                    val = value.Value,
                };
            }
            return result;
        }

        private static E2009.CT_FontScheme ToExcel2009(FontScheme value)
        {
            E2009.CT_FontScheme result = null;
            if (value != null)
            {
                result = new E2009.CT_FontScheme
                {
                    val = ToExcel2009(value.Value),
                };
            }
            return result;
        }

        private static E2009.CT_FontSize ToExcel2009(FontSize value)
        {
            E2009.CT_FontSize result = null;
            if (value != null)
            {
                result = new E2009.CT_FontSize
                {
                    val = (value.Value),
                };
            }
            return result;
        }

        private static E2009.CT_UnderlineProperty ToExcel2009(UnderlineProperty value)
        {
            E2009.CT_UnderlineProperty result = null;
            if (value != null)
            {
                result = new E2009.CT_UnderlineProperty
                {
                    val = ToExcel2009(value.Value),
                };
            }
            return result;
        }

        private static E2009.CT_VerticalAlignFontProperty ToExcel2009(VerticalAlignFontProperty value)
        {
            E2009.CT_VerticalAlignFontProperty result = null;
            if (value != null)
            {
                result = new E2009.CT_VerticalAlignFontProperty
                {
                    val = ToExcel2009(value.Value),
                };
            }
            return result;
        }

        private static E2009.CT_BorderPr ToExcel2009(BorderPr value)
        {
            E2009.CT_BorderPr result = null;
            if (value != null)
            {
                result = new E2009.CT_BorderPr
                {
                    color = ToExcel2009(value.Color),
                    style = ToExcel2009(value.Style??BorderStyle.None),
                };
            }
            return result;
        }

        private static E2009.CT_PatternFill ToExcel2009(PatternFill value)
        {
            E2009.CT_PatternFill result = null;
            if (value != null)
            {
                bool patternTypeSpecified;
                var patternType = ToExcel2009(value.PatternType, out patternTypeSpecified);
                result = new E2009.CT_PatternFill
                {
                    patternType = patternType,
                    patternTypeSpecified = patternTypeSpecified,
                    bgColor = ToExcel2009(value.BackgroundColor),
                    fgColor = ToExcel2009(value.ForegroundColor),

                };
            }
            return result;
        }

        private static E2009.CT_GradientFill ToExcel2009(GradientFill value)
        {
            E2009.CT_GradientFill result = null;
            if (value != null)
            {
                result = new E2009.CT_GradientFill
                {
                 type = ToExcel2009(value.GradientType),
                 bottom = value.Bottom,
                 degree = value.Degree,
                 left =  value.Left,
                 right = value.Right,
                 top = value.Top,
                };
                {
                    var list = new List<E2009.CT_GradientStop>();
                    if (value.GradientStop != null)
                        foreach (var i in value.GradientStop)
                        {
                            var item = ToExcel2009(i);
                            if (item != null)
                            {
                                list.Add(item);
                            }
                        }
                    result.stop = list.ToArray();
                }
            }
            return result;
        }

        private static E2009.ST_HorizontalAlignment ToExcel2009(HorizontalAlignment? value, out bool specified)
        {
            specified = value != null;
            switch (value)
            {
                case HorizontalAlignment.General:
                    return E2009.ST_HorizontalAlignment.general;
                case HorizontalAlignment.Left:
                    return E2009.ST_HorizontalAlignment.left;
                case HorizontalAlignment.Center:
                    return E2009.ST_HorizontalAlignment.center;
                case HorizontalAlignment.Right:
                    return E2009.ST_HorizontalAlignment.right;
                case HorizontalAlignment.Fill:
                    return E2009.ST_HorizontalAlignment.fill;
                case HorizontalAlignment.Justify:
                    return E2009.ST_HorizontalAlignment.justify;
                case HorizontalAlignment.CenterContinuous:
                    return E2009.ST_HorizontalAlignment.centerContinuous;
                case HorizontalAlignment.Distributed:
                    return E2009.ST_HorizontalAlignment.distributed;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.ST_VerticalAlignment ToExcel2009(VerticalAlignment value)
        {
            switch (value)
            {
                case VerticalAlignment.Top:
                    return E2009.ST_VerticalAlignment.top;
                case VerticalAlignment.Center:
                    return E2009.ST_VerticalAlignment.center;
                case VerticalAlignment.Bottom:
                    return E2009.ST_VerticalAlignment.bottom;
                case VerticalAlignment.Justify:
                    return E2009.ST_VerticalAlignment.justify;
                case VerticalAlignment.Distributed:
                    return E2009.ST_VerticalAlignment.distributed;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.ST_CfvoType ToExcel2009(ConditionalFormatValueObjectType value)
        {
            switch (value)
            {
                case ConditionalFormatValueObjectType.Formula:
                    return E2009.ST_CfvoType.formula;
                case ConditionalFormatValueObjectType.Max:
                    return E2009.ST_CfvoType.max;
                case ConditionalFormatValueObjectType.Min:
                    return E2009.ST_CfvoType.min;
                case ConditionalFormatValueObjectType.Num:
                    return E2009.ST_CfvoType.num;
                case ConditionalFormatValueObjectType.Percent:
                    return E2009.ST_CfvoType.percent;
                case ConditionalFormatValueObjectType.Percentile:
                    return E2009.ST_CfvoType.percentile;
                case ConditionalFormatValueObjectType.AutoMin:
                    return E2009.ST_CfvoType.autoMin;
                case ConditionalFormatValueObjectType.AutoMax:
                    return E2009.ST_CfvoType.autoMax;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static byte[] ToExcel2009ArgbToRgb(Argb value)
        {
            if (value?.Value != null && (value.Value.Length == 4 || value.Value.Length == 3))
            {
                byte[] result = new byte[value.Value.Length];
                Array.Copy(value.Value, result, value.Value.Length);
                return result;
            }
            //TODO ! WARNING ! такой ситуациине должно быть разобраться как она получается
            // файл qwe.xlsx
            return new byte[4];
        }

        private static E2009.ST_FontScheme ToExcel2009(FontSchemeEnum value)
        {
            switch (value)
            {
                case FontSchemeEnum.None:
                    return E2009.ST_FontScheme.none;
                case FontSchemeEnum.Minor:
                    return E2009.ST_FontScheme.minor;
                case FontSchemeEnum.Major:
                    return E2009.ST_FontScheme.major;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.ST_UnderlineValues ToExcel2009(UnderlineValues value)
        {
            switch (value)
            {
                case UnderlineValues.Single:
                    return E2009.ST_UnderlineValues.single;
                case UnderlineValues.Double:
                    return  E2009.ST_UnderlineValues.@double;
                case UnderlineValues.SingleAccounting:
                    return E2009.ST_UnderlineValues.singleAccounting;
                case UnderlineValues.DoubleAccounting:
                    return E2009.ST_UnderlineValues.doubleAccounting;
                case UnderlineValues.None:
                    return E2009.ST_UnderlineValues.none;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.ST_VerticalAlignRun ToExcel2009(VerticalAlignRun value)
        {
            switch (value)
            {
                case VerticalAlignRun.Baseline:
                    return E2009.ST_VerticalAlignRun.baseline;
                case VerticalAlignRun.Superscript:
                    return E2009.ST_VerticalAlignRun.superscript;
                case VerticalAlignRun.Subscript:
                    return E2009.ST_VerticalAlignRun.subscript;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.ST_BorderStyle ToExcel2009(BorderStyle value)
        {
            switch (value)
            {
                case BorderStyle.None:
                    return E2009.ST_BorderStyle.none;
                case BorderStyle.Thin:
                    return E2009.ST_BorderStyle.thin;
                case BorderStyle.Medium:
                    return E2009.ST_BorderStyle.medium;
                case BorderStyle.Dashed:
                    return E2009.ST_BorderStyle.dashed;
                case BorderStyle.Dotted:
                    return E2009.ST_BorderStyle.dotted;
                case BorderStyle.Thick:
                    return E2009.ST_BorderStyle.thick;
                case BorderStyle.Double:
                    return E2009.ST_BorderStyle.@double;
                case BorderStyle.Hair:
                    return E2009.ST_BorderStyle.hair;
                case BorderStyle.MediumDashed:
                    return E2009.ST_BorderStyle.mediumDashed;
                case BorderStyle.DashDot:
                    return E2009.ST_BorderStyle.dashDot;
                case BorderStyle.MediumDashDot:
                    return E2009.ST_BorderStyle.mediumDashDot;
                case BorderStyle.DashDotDot:
                    return E2009.ST_BorderStyle.dashDotDot;
                case BorderStyle.MediumDashDotDot:
                    return E2009.ST_BorderStyle.mediumDashDotDot;
                case BorderStyle.SlantDashDot:
                    return E2009.ST_BorderStyle.slantDashDot;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.ST_PatternType ToExcel2009(PatternType? value, out bool specified)
        {
            specified = value != null;
            switch (value)
            {
                case PatternType.None:
                    return  E2009.ST_PatternType.none;
                case PatternType.Solid:
                    return E2009.ST_PatternType.solid;
                case PatternType.MediumGray:
                    return E2009.ST_PatternType.mediumGray;
                case PatternType.DarkGray:
                    return E2009.ST_PatternType.darkGray;
                case PatternType.LightGray:
                    return E2009.ST_PatternType.lightGray;
                case PatternType.DarkHorizontal:
                    return E2009.ST_PatternType.darkHorizontal;
                case PatternType.DarkVertical:
                    return E2009.ST_PatternType.darkVertical;
                case PatternType.DarkDown:
                    return E2009.ST_PatternType.darkDown;
                case PatternType.DarkUp:
                    return E2009.ST_PatternType.darkUp;
                case PatternType.DarkGrid:
                    return E2009.ST_PatternType.darkGrid;
                case PatternType.DarkTrellis:
                    return E2009.ST_PatternType.darkTrellis;
                case PatternType.LightHorizontal:
                    return E2009.ST_PatternType.lightHorizontal;
                case PatternType.LightVertical:
                    return E2009.ST_PatternType.lightVertical;
                case PatternType.LightDown:
                    return E2009.ST_PatternType.lightDown;
                case PatternType.LightUp:
                    return E2009.ST_PatternType.lightUp;
                case PatternType.LightGrid:
                    return E2009.ST_PatternType.lightGrid;
                case PatternType.LightTrellis:
                    return E2009.ST_PatternType.lightTrellis;
                case PatternType.Gray125:
                    return E2009.ST_PatternType.gray125;
                case PatternType.Gray0625:
                    return E2009.ST_PatternType.gray0625;
                case null:
                    return E2009.ST_PatternType.none;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.ST_GradientType ToExcel2009(GradientType value)
        {
            switch (value)
            {
                case GradientType.Linear:
                    return E2009.ST_GradientType.linear;
                case GradientType.Path:
                    return E2009.ST_GradientType.path;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        private static E2009.CT_GradientStop ToExcel2009(GradientStop value)
        {
            E2009.CT_GradientStop result = null;
            if (value != null)
            {
                result = new E2009.CT_GradientStop
                {
                    color = ToExcel2009(value.Color),
                    position =  value.Position,
                };
                
            }
            return result;

        }
        #endregion

        #region Substitution

        public static void Subsitute(ConditionalFormatRule cf2006From, ConditionalFormatRule cf2009To)
        {
            if (cf2009To.Priority == null)
            {
                cf2009To.Priority = cf2006From.Priority;
            }
            if (cf2009To.DataBar != null && cf2009To.DataBar.FillColor == null)
            {
                cf2009To.DataBar.FillColor = cf2006From?.DataBar?.FillColor;
            }
        }

        #endregion
    }

    public static class CopyConditionalFormatsHelper
    {
        public static void CopyConditionalFormatting(ReoGrid.Worksheet source, Cell src, Cell dst)
        {
            dst.Worksheet.ConditionalFormats = dst.Worksheet.ConditionalFormats ?? new List<ConditionalFormat>();
            var conditionalFormatsSrc = source.ConditionalFormats;
            var conditionalFormatting = conditionalFormatsSrc?.Where(arg => AddressInRangePosition(src.Address, arg.Sqref.Text)).ToArray();
            if (conditionalFormatting != null)
                foreach (var cf in conditionalFormatting)
                {
                    var conditionalFormat = new ConditionalFormat
                    {
                        Pivot = cf.Pivot,
                        Sqref = new Sqref { Text = dst.Address }
                    };
                    foreach (var rule in cf.Rules)
                    {
                        var newRule = (ConditionalFormatRule)rule.Clone();
                        newRule.Formula.Clear();
                        foreach (var formula in rule.Formula)
                            newRule.Formula.Add(new FormulaItem
                            {
                                Value = CellAddressRegex.Replace(formula.Value, m =>
                                {
                                    var address = m.Value;
                                    if (!CellPosition.IsValidAddress(address)) return address;
                                    var position = new CellPosition(address);
                                    position.Col += position.ColumnProperty == PositionProperty.Absolute ? 0 : dst.Column - src.Column;
                                    position.Row += position.RowProperty == PositionProperty.Absolute ? 0 : dst.Row - src.Row;
                                    return position.ToString();
                                })
                            });
                        conditionalFormat.Rules.Add(newRule);
                    }
                    dst.Worksheet.ConditionalFormats.Add(conditionalFormat);
                }
        }

        private static bool AddressInRangePosition(string address, string range)
        {
            var rangePos = new RangePosition(range);
            return rangePos.Contains(new CellPosition(address));
        }

        private static readonly Regex CellAddressRegex = new Regex(@"\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?");
    }
}
