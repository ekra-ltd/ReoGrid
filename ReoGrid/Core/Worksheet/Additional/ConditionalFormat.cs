using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using unvell.ReoGrid.Formula;

namespace unvell.ReoGrid.Core.Worksheet.Additional
{
	public class ConditionalFormat
	{
		public List<ConditionalFormatRule> Rules { get; set; } = new List<ConditionalFormatRule>();

		public Sqref Sqref { get; set; }

		public bool Pivot { get; set; }

	}

	public class Sqref
	{
		public bool? Edited { get; set; }
		public bool? Split { get; set; }
		public bool? Adjusted { get; set; }
		public bool? Adjust { get; set; }
		public string Text { get; set; }
	}


	public class FormulaItem
	{
		public string Value { get; set; }

		internal STNode FormulaTree { get; set; }
	}

	public class ConditionalFormatRule : ICloneable
	{
		/// <summary>
		/// &lt;xsd:element ref="xm:f" minOccurs="0" maxOccurs="3"/&gt;
		/// </summary>
		public List<FormulaItem> Formula { get; set; } = new List<FormulaItem>();

		public ColorScale ColorScale { get; set; }

		/// <summary>
		/// &lt;xsd:element name="dataBar" type="CT_DataBar" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public DataBar DataBar { get; set; }

		/// <summary>
		/// &lt;xsd:element name="iconSet" type="CT_IconSet" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public IconSet IconSet { get; set; }

		/// <summary>
		/// &lt;xsd:element name="dxf" type="x:CT_Dxf" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public DifferentialFormat DifferentialFormat { get; set; }

		//<xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>

		/// <summary>
		/// &lt;xsd:attribute name="type" type="x:ST_CfType" use="optional"/&gt;
		/// </summary>
		public ConditionalFormatType? Type { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="priority" type="xsd:int" use="optional"/&gt;
		/// </summary>
		public int? Priority { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="stopIfTrue" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? StopIfTrue { get; set; }


		/// <summary>
		/// &lt;xsd:attribute name="aboveAverage" type="xsd:boolean" use="optional" default="true"/&gt;
		/// </summary>
		public bool? AboveAverage { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="percent" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? Percent { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="Bottom" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? Bottom { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="operator" type="x:ST_ConditionalFormattingOperator" use="optional"/&gt;
		/// </summary>
		public ConditionalFormattingOperator? Operator { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="text" type="xsd:string" use="optional"/&gt;
		/// </summary>
		public string Text { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="timePeriod" type="x:ST_TimePeriod" use="optional"/&gt;
		/// </summary>
		public TimePeriod? TimePeriod { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="rank" type="xsd:unsignedInt" use="optional"/&gt;
		/// </summary>
		public uint? Rank { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="stdDev" type="xsd:int" use="optional"/&gt;
		/// </summary>
		public int? StdDev { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="equalAverage" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? EqualAverage { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="activePresent" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? ActivePercent { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="id" type="x:ST_Guid" use="optional"/&gt;
		/// </summary>
		public string SGuid { get; set; }

		[XmlIgnore]
		public string Ext2009Id { get; set; }

		public object Clone()
			=> new ConditionalFormatRule
			{
				Formula = new List<FormulaItem>(Formula),
				ColorScale = ColorScale,
				DataBar = DataBar,
				IconSet = IconSet,
				DifferentialFormat = DifferentialFormat,
				Type = Type,
				Priority = Priority,
				StopIfTrue = StopIfTrue,
				AboveAverage = AboveAverage,
				Percent = Percent,
				Bottom = Bottom,
				Operator = Operator,
				Text = Text,
				TimePeriod = TimePeriod,
				Rank = Rank,
				StdDev = StdDev,
				EqualAverage = EqualAverage,
				ActivePercent = ActivePercent,
				SGuid = SGuid,
				Ext2009Id = Ext2009Id
			};
	}

	/// <summary>
	/// Describes the values of the interpolation points in a gradient scale.
	/// </summary>
	public class ConditionalFormatValueObject
	{
		public ConditionalFormatValueObjectType Type { get; set; }

		/// <summary>
		/// Greater Than Or Equal.
		/// For icon sets, determines whether this threshold value uses the greater than or equal to
		/// operator. 0 indicates 'greater than' is used instead of 'greater than or equal to'.
		/// </summary>
		public bool Gte { get; set; } = true;

		public FormulaItem Formula { get; set; }

		//<xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
	}

	/// <summary>
	/// This simple type expresses the type of the conditional formatting value object (cfvo). 
	/// In general the cfvo specifies one value used in the gradated scale(max, min, midpoint, etc).
	/// </summary>
	public enum ConditionalFormatValueObjectType
	{
		/// <summary>
		/// The minimum/ midpoint / maximum value for the gradient is determined by a formula.
		/// </summary>
		Formula,

		/// <summary>
		/// Indicates that the maximum value in the range shall be used as the maximum value for the gradient.
		/// </summary>
		Max,

		/// <summary>
		/// Indicates that the minimum value in the range shall be used as the minimum value for the gradient.
		/// </summary>
		Min,

		/// <summary>
		/// Indicates that the minimum / midpoint / maximum value for the gradient is specified by a constant numeric value.
		/// </summary>
		Num,

		/// <summary>
		/// Value indicates a percentage between the minimum and maximum values in the range shall be used as the minimum / midpoint / maximum value for the gradient.
		/// </summary>
		Percent,

		/// <summary>
		/// Value indicates a percentile ranking in the range shall be used as the minimum / midpoint / maximum value for the gradient.
		/// </summary>
		Percentile,

		/// <summary>
		/// Обьявлено в http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
		/// </summary>
		AutoMin,

		/// <summary>
		/// Обьявлено в http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
		/// </summary>
		AutoMax,
	}

	/// <summary>
	/// Describes a gradated color scale in this conditional formatting rule.
	/// </summary>
	public class ColorScale
	{
		/// <summary>
		/// &lt;xsd:element name="cfvo" type="CT_Cfvo" minOccurs="2" maxOccurs="unbounded"/&gt;
		/// </summary>
		public List<ConditionalFormatValueObject> CondittionalFormatValue { get; set; } =
			new List<ConditionalFormatValueObject>();

		/// <summary>
		/// &lt;xsd:element name = "color" type="x:CT_Color" minOccurs="2" maxOccurs="unbounded"/&gt;
		/// </summary>
		public List<Color> Color { get; set; } = new List<Color>();
	}

	/// <summary>
	/// One of the colors associated with the data bar or color scale.
	/// The auto attribute shall not be used in the context of data bars.
	/// </summary>
	public class Color
	{
		/// <summary>
		/// A boolean value indicating the color is automatic and system color dependent. 
		/// </summary>
		public bool? Automatic { get; set; }

		/// <summary>
		/// Indexed color value. Only used for backwards compatibility. References a color in indexedColors.
		/// </summary>
		public uint? Indexed { get; set; }

		/// <summary>
		/// Standard Alpha Red Green Blue color value (ARGB).
		/// </summary>
		public Argb RgbColorValue { get; set; }

		/// <summary>
		/// A zero-based index into the &lt;clrScheme&gt collection (§20.1.6.2), referencing a particular &lt;sysClr&gt or&lt;srgbClr&gt value expressed in the Theme part.
		/// </summary>
		public uint? ThemeColor { get; set; }

		/// <summary>
		/// Specifies the tint value applied to the color.
		/// If tint is supplied, then it is applied to the RGB value of the color to determine the final color applied.
		/// The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and 1.0 means 100% lighten. Also, 0.0 means no change.
		/// In loading the RGB value, it is converted to HLS where HLS values are (0..HLSMAX), where HLSMAX is currently 255.
		/// </summary>
		public double? TInt { get; set; }
	}

	public class Argb
	{
		public byte[] Value { get; set; } = new byte[4];
		
		public Graphics.SolidColor ToSolidColor()
		{
			return new Graphics.SolidColor(Value[0], Value[1], Value[2], Value[3]);
		}
	}

	/// <summary>
	/// Describes a data bar conditional formatting rule.
	/// </summary>
	public class DataBar
	{
		/// <summary>
		/// &lt;xsd:element name="cfvo" type="CT_Cfvo" minOccurs="2" maxOccurs="2"/&gt;
		/// </summary>
		public List<ConditionalFormatValueObject> CondittionalFormatValue { get; set; } =
			new List<ConditionalFormatValueObject>();

		/// <summary>
		/// &lt;xsd:element name="fillColor" type="x:CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color FillColor { get; set; }

		/// <summary>
		/// &lt;xsd:element name="borderColor" type="x:CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color BorderColor { get; set; }

		/// <summary>
		/// &lt;xsd:element name="negativeFillColor" type="x:CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color NegativeFillColor { get; set; }

		/// <summary>
		/// &lt;xsd:element name="negativeBorderColor" type="x:CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color NegativeBorderColor { get; set; }

		/// <summary>
		/// &lt;xsd:element name="axisColor" type="x:CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color AxisColor { get; set; }

		/// <summary>
		/// The minimum length of the data bar, as a percentage of the cell width.
		/// &lt;xsd:attribute name="minLength" type="xsd:unsignedInt" use="optional" default="10"/&gt;
		/// </summary>
		public uint? MinLength { get; set; }

		/// <summary>
		/// The maximum length of the data bar, as a percentage of the cell width.
		/// &lt;xsd:attribute name="maxLength" type="xsd:unsignedInt" use="optional" default="90"/&gt;
		/// </summary>
		public uint? MaxLength { get; set; }

		/// <summary>
		/// Indicates whether to show the values of the cells on which this data bar is applied.
		/// &lt;xsd:attribute name="showValue" type="xsd:boolean" use="optional" default="true"/&gt;
		/// </summary>
		public bool? ShowValue { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="border" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? Border { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="gradient" type="xsd:boolean" use="optional" default="true"/&gt;
		/// </summary>
		public bool? Gradient { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="direction" type="ST_DataBarDirection" use="optional" default="context"/&gt;
		/// </summary>
		public DataBarDirection? Direction { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="negativeBarColorSameAsPositive" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? NegativeBarColorSameAsPositive { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="negativeBarBorderColorSameAsPositive" type="xsd:boolean" use="optional" default="true"/&gt;
		/// </summary>
		public bool? NegativeBarBorderColorSameAsPositive { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="axisPosition" type="ST_DataBarAxisPosition" use="optional" default="automatic"/&gt;
		/// </summary>
		public DatabarAxisPosition? AxisPosition { get; set; }
	}

	public enum DataBarDirection
	{
		Context,
		LeftToRight,
		RightToleft,
	}

	public enum DatabarAxisPosition
	{
		Automatic,
		Middle,
		None,
	}

	/// <summary>
	/// Describes an icon set conditional formatting rule.
	/// </summary>
	public class IconSet
	{
		/// <summary>
		/// &lt;xsd:element name="cfvo" type="CT_Cfvo" minOccurs="2" maxOccurs="unbounded"/&gt;
		/// </summary>
		public List<ConditionalFormatValueObject> CondittionalFormatValue { get; set; } =
			new List<ConditionalFormatValueObject>();

		/// <summary>
		/// &lt;xsd:element name="cfIcon" type="CT_CfIcon" minOccurs="0" maxOccurs="5"/&gt;
		/// </summary>
		public List<CfIcon> Cficon { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="iconSet" type="ST_IconSetType" use="optional" default="3TrafficLights1"/&gt;
		/// </summary>
		public IconSetType IconSetType { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="showValue" type="xsd:boolean" use="optional" default="true"/&gt;
		/// </summary>
		public bool? ShowValues { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="percent" type="xsd:boolean" use="optional" default="true"/&gt;
		/// </summary>
		public bool? Percent { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="reverse" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? Reverse { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="custom" type="xsd:boolean" use="optional" default="false"/&gt;
		/// </summary>
		public bool? Custom { get; set; }
	}

	public class CfIcon
	{
		/// <summary>
		/// &lt;xsd:attribute name="iconSet" type="ST_IconSetType" use="required"/&gt;
		/// </summary>
		public IconSetType IconSet { get; set; }


		/// <summary>
		/// &lt;xsd:attribute name="iconId" type="xsd:unsignedInt" use="required"/&gt;
		/// </summary>
		public uint IconId { get; set; }
	}

	public enum IconSetType
	{
		IconSet_3Arrows,
		IconSet_3ArrowsGray,
		IconSet_3Flags,
		IconSet_3TrafficLights1,
		IconSet_3TrafficLights2,
		IconSet_3Signs,
		IconSet_3Symbols,
		IconSet_3Symbols2,
		IconSet_4Arrows,
		IconSet_4ArrowsGray,
		IconSet_4RedToBlack,
		IconSet_4Rating,
		IconSet_4TrafficLights,
		IconSet_5Arrows,
		IconSet_5ArrowsGray,
		IconSet_5Rating,
		IconSet_5Quarters,
		IconSet_3Stars,
		IconSet_3Triangles,
		IconSet_5Boxes,
		IconSet_NoIcons,
	}

	public class DifferentialFormat
	{
		/// <summary>
		/// &lt;xsd:element name="font" type="CT_Font" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Font Font { get; set; }

		/// <summary>
		/// &lt;xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public NumberFormat NumberFormat { get; set; }

		/// <summary>
		/// &lt;xsd:element name="fill" type="CT_Fill" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Fill Fill { get; set; }

		/// <summary>
		/// &lt;xsd:element name="alignment" type="CT_CellAlignment" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public CellAlignment CellAlignment { get; set; }

		/// <summary>
		/// &lt;xsd:element name="border" type="CT_Border" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Border Border { get; set; }

		/// <summary>
		/// &lt;xsd:element name="protection" type="CT_CellProtection" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public CellProtection CellProtection { get; set; }

		//<xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
	}

	public class Font
	{
		/// <summary>
		/// &lt;xsd:element name="name" type="CT_FontName" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public FontName Name { get; set; }

		/// <summary>
		/// &lt;xsd:element name="charset" type="CT_IntProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public IntProperty Charset { get; set; }

		/// <summary>
		/// &lt;xsd:element name="family" type="CT_FontFamily" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public FontFamily Family { get; set; }

		/// <summary>
		/// &lt;xsd:element name="b" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BooleanProperty Bold { get; set; }

		/// <summary>
		/// &lt;xsd:element name="i" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BooleanProperty Italic { get; set; }

		/// <summary>
		/// &lt;xsd:element name="strike" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BooleanProperty Strike { get; set; }

		/// <summary>
		/// &lt;xsd:element name="outline" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BooleanProperty Outline { get; set; }


		/// <summary>
		/// &lt;xsd:element name="shadow" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BooleanProperty Shadow { get; set; }

		/// <summary>
		/// &lt;xsd:element name="condense" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BooleanProperty Condense { get; set; }

		/// <summary>
		/// &lt;xsd:element name="extend" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BooleanProperty Extend { get; set; }

		/// <summary>
		/// &lt;xsd:element name="color" type="CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color Color { get; set; }

		/// <summary>
		/// &lt;xsd:element name="sz" type="CT_FontSize" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public FontSize FontSize { get; set; }

		/// <summary>
		/// &lt;xsd:element name="u" type="CT_UnderlineProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public UnderlineProperty Underline { get; set; }

		/// <summary>
		/// &lt;xsd:element name="vertAlign" type="CT_VerticalAlignFontProperty" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public VerticalAlignFontProperty VerticalAlign { get; set; }

		/// <summary>
		/// &lt;xsd:element name="scheme" type="CT_FontScheme" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public FontScheme FontScheme { get; set; }
	}

	public class FontName
	{
		public string Value { get; set; }
	}

	public class IntProperty
	{
		public int Value { get; set; }
	}

	public class FontFamily
	{
		public string Value { get; set; }
	}

	public class BooleanProperty
	{
		public bool Value { get; set; }
	}

	public class FontSize
	{
		public double Value { get; set; }
	}

	public class UnderlineProperty
	{
		public UnderlineValues Value { get; set; }
	}

	public enum UnderlineValues
	{
		Single,
		Double,
		SingleAccounting,
		DoubleAccounting,
		None,
	}

	public class VerticalAlignFontProperty
	{
		public VerticalAlignRun Value { get; set; }
	}

	public enum VerticalAlignRun
	{
		Baseline,
		Superscript,
		Subscript,
	}

	public class FontScheme
	{
		public FontSchemeEnum Value { get; set; }
	}

	public enum FontSchemeEnum
	{
		None,
		Minor,
		Major,
	}

	/// <summary>
	/// This element specifies number format properties which indicate how to format and render the numeric value of a cell.
	/// </summary>
	public class NumberFormat
	{
		/// <summary>
		/// &lt;xsd:attribute name="numFmtId" type="ST_NumFmtId" use="required"/&gt;
		/// </summary>
		public uint NumberFormatId { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="formatCode" type="s:ST_Xstring" use="required"/&gt;
		/// </summary>
		public string FormatCode { get; set; }
	}

	public class Fill
	{
		/// <summary>
		/// &lt;xsd:element name="patternFill" type="CT_PatternFill" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public PatternFill PatternFill { get; set; }


		/// <summary>
		/// &lt;xsd:element name="gradientFill" type="CT_GradientFill" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public GradientFill GradientFill { get; set; }
	}

	public class PatternFill
	{
		/// <summary>
		/// &lt;xsd:element name="fgColor" type="CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color ForegroundColor { get; set; }

		/// <summary>
		/// &lt;xsd:element name="bgColor" type="CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color BackgroundColor { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="patternType" type="ST_PatternType" use="optional"/&gt;
		/// </summary>
		public PatternType? PatternType { get; set; }
	}

	public enum PatternType
	{
		None,
		Solid,
		MediumGray,
		DarkGray,
		LightGray,
		DarkHorizontal,
		DarkVertical,
		DarkDown,
		DarkUp,
		DarkGrid,
		DarkTrellis,
		LightHorizontal,
		LightVertical,
		LightDown,
		LightUp,
		LightGrid,
		LightTrellis,
		Gray125,
		Gray0625,
	}

	public class GradientFill
	{
		/// <summary>
		/// &lt;xsd:element name="stop" type="CT_GradientStop" minOccurs="0" maxOccurs="unbounded"/&gt;
		/// </summary>
		public List<GradientStop> GradientStop { get; set; }= new List<GradientStop>();

		/// <summary>
		/// &lt;xsd:attribute name="type" type="ST_GradientType" use="optional" default="linear"/&gt;
		/// </summary>
		public GradientType GradientType { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="degree" type="xsd:double" use="optional" default="0"/&gt;
		/// </summary>
		public double Degree { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="left" type="xsd:double" use="optional" default="0"/&gt;
		/// </summary>
		public double Left { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="right" type="xsd:double" use="optional" default="0"/&gt;
		/// </summary>
		public double Right { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="top" type="xsd:double" use="optional" default="0"/&gt;
		/// </summary>
		public double Top { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="Bottom" type="xsd:double" use="optional" default="0"/&gt;
		/// </summary>
		public double Bottom { get; set; }
	}

	public class GradientStop
	{
		/// <summary>
		/// &lt;xsd:element name="color" type="CT_Color" minOccurs="1" maxOccurs="1"/&gt;
		/// </summary>
		public Color Color { get; set; }
		/// <summary>
		/// &lt;xsd:attribute name="position" type="xsd:double" use="required"/&gt;
		/// </summary>
		public double Position { get; set; }
	}

	public enum GradientType
	{
		Linear,
		Path,
	}

	public class CellAlignment
	{
		/// <summary>
		/// &lt;xsd:attribute name="horizontal" type="ST_HorizontalAlignment" use="optional"/&gt;
		/// </summary>
		public HorizontalAlignment? Horizontal { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="vertical" type="ST_VerticalAlignment" default="Bottom" use="optional"/&gt;
		/// </summary>
		public VerticalAlignment Vertical { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="textRotation" type="ST_TextRotation" use="optional"/&gt;
		/// непонятное мне преобразование - пусть будет строкой
		/// </summary>
		public string TextRotation { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="wrapText" type="xsd:boolean" use="optional"/&gt;
		/// </summary>
		public bool? WrapText { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="indent" type="xsd:unsignedInt" use="optional"/&gt;
		/// </summary>
		public uint? Indent { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="relativeIndent" type="xsd:int" use="optional"/&gt;
		/// </summary>
		public int? RelativeIndent { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="justifyLastLine" type="xsd:boolean" use="optional"/&gt;
		/// </summary>
		public bool? JustifyLastLine { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="shrinkToFit" type="xsd:boolean" use="optional"/&gt;
		/// </summary>
		public bool? ShrinkToFit { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="readingOrder" type="xsd:unsignedInt" use="optional"/&gt;
		/// </summary>
		public uint? ReadingOrder { get; set; }
	}

	public enum HorizontalAlignment
	{
		General,
		Left,
		Center,
		Right,
		Fill,
		Justify,
		CenterContinuous,
		Distributed,
	}

	public enum VerticalAlignment
	{
		Top,
		Center,
		Bottom,
		Justify,
		Distributed,
	}

	public class Border
	{
		/// <summary>
		/// &lt;xsd:element name="start" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BorderPr Start { get; set; }
		/// <summary>
		/// &lt;xsd:element name="end" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BorderPr End { get; set; }
		/// <summary>
		/// &lt;xsd:element name="top" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BorderPr Top{ get; set; }
		/// <summary>
		/// &lt;xsd:element name="Bottom" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BorderPr Bottom { get; set; }
		/// <summary>
		/// &lt;xsd:element name="diagonal" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BorderPr Diagonal { get; set; }
		/// <summary>
		/// &lt;xsd:element name="vertical" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BorderPr Vertical { get; set; }
		/// <summary>
		/// &lt;xsd:element name="horizontal" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public BorderPr Horizontal { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="diagonalUp" type="xsd:boolean" use="optional"/&gt;
		/// </summary>
		public bool? DiagonalUp { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="diagonalDown" type="xsd:boolean" use="optional"/&gt;
		/// </summary>
		public bool? DiagonalDown { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="outline" type="xsd:boolean" use="optional" default="true"/&gt;
		/// </summary>
		public bool? Outline { get; set; }
	}

	public class BorderPr
	{
		/// <summary>
		/// &lt;xsd:element name="color" type="CT_Color" minOccurs="0" maxOccurs="1"/&gt;
		/// </summary>
		public Color Color { get; set; }

		/// <summary>
		/// &lt;xsd:attribute name="style" type="ST_BorderStyle" use="optional" default="none"/&gt;
		/// </summary>
		public BorderStyle? Style { get; set; }
	}

	public enum BorderStyle
	{
		None,
		Thin,
		Medium,
		Dashed,
		Dotted,
		Thick,
		Double,
		Hair,
		MediumDashed,
		DashDot,
		MediumDashDot,
		DashDotDot,
		MediumDashDotDot,
		SlantDashDot,
	}

	public class CellProtection
	{
		/// <summary>
		/// &lt;xsd:attribute name="locked" type="xsd:boolean" use="optional"/&gt;
		/// </summary>
		public bool? Locked { get; set; }
		/// <summary>
		/// &lt;xsd:attribute name="hidden" type="xsd:boolean" use="optional"/&gt;
		/// </summary>
		public bool? Hidden { get; set; }
	}

	public enum ConditionalFormatType
	{
		Expression,
		CellIs,
		ColorScale,
		DataBar,
		IconSet,
		Top10,
		UniqueValues,
		DuplicateValues,
		ContainsText,
		NotContainsText,
		BeginsWith,
		EndsWith,
		ContainsBlanks,
		NotContainsBlanks,
		ContainsErrors,
		NotContainsErrors,
		TimePeriod,
		AboveAverage,
	}

	public enum ConditionalFormattingOperator
	{
		LessThan,
		LessThanOrEqual,
		Equal,
		NotEqual,
		GreaterThanOrEqual,
		GreaterThan,
		Between,
		NotBetween,
		ContainsText,
		NotContains,
		BeginsWith,
		EndsWith,
	}

	public enum TimePeriod
	{
		Today,
		Yesterday,
		Tomorrow,
		Last7Days,
		ThisMonth,
		LastMonth,
		NextMonth,
		ThisWeek,
		LastWeek,
		NextWeek,
	}
}
