/*****************************************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Jingwood <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2025 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2025 UNVELL Inc. All rights reserved.
 * 
 ****************************************************************************/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using unvell.ReoGrid.Core;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Utility;
using System.Text.RegularExpressions;

namespace unvell.ReoGrid.DataFormat
{
	/// <summary>
	/// Datetime data formatter
	/// </summary>
	public class DateTimeDataFormatter : IDataFormatter
	{
		public object GetDefaultDataFormatArgs()
		{
			return new DateTimeFormatArgs()
			{
				Format = @"dd/MM/yyyy HH:mm:ss",
				CultureName = @"ru-RU"
			};
		}
		private class DateTimeFormats
		{
			public DateTimeFormats(CultureInfo culture)
			{
				Formats = culture.DateTimeFormat.GetAllDateTimePatterns().Distinct().ToArray();
			}

			public string[] Formats { get; }
		}

		private static DateTimeFormats CurrentCultureDateTimeFormats { get; } = new DateTimeFormats(Thread.CurrentThread.CurrentCulture);

		private static DateTimeFormats InvariantCultureDateTimeFormats { get; } = new DateTimeFormats(CultureInfo.InvariantCulture);

		/// <summary>
		/// Список правил отображения, для которых нужно выставлять формат "Дата/время"
		/// </summary>
		private static string[] _dateTimeViewRules =
		{
			"%REPORT.DATE%",         @"%ОТЧЕТ.ДАТА%",
			"%REPORT.DATEOFREPORT%", @"%ОТЧЕТ.ДАТАФОРМИРОВАНИЯ%",
			"%REPORT.TIMEOFREPORT%", @"%ОТЧЕТ.ВРЕМЯФОРМИРОВАНИЯ%"
		};

		/// <summary>
		/// Флаг изменения списка правил отображения, для которых нужно выставлять формат "Дата/время"
		/// </summary>
		private static volatile bool _dateTimeViewRulesChanged;

		/// <summary>
		/// Список заданных правил отображения, для которых нужно выставлять формат "Дата/время"
		/// </summary>
		// ReSharper disable once MemberCanBePrivate.Global
		public static string[] DateTimeViewRules
		{
			get => _dateTimeViewRules;
			// ReSharper disable once UnusedMember.Global
			set
			{
				_dateTimeViewRules = value;
				_dateTimeViewRulesChanged = true;
			}
		}

		private static Regex _hasDateTimeViewRuleRegex = null;

		private static bool TryParseDateTime(DateTimeFormats data, string src, out DateTime value)
		{
			var result = Constants.ExcelZeroDatePoint;
			var parseResult = data.Formats.Any(format => DateTime.TryParseExact(src, format, Thread.CurrentThread.CurrentCulture, DateTimeStyles.None, out result));
			value = result;
			return parseResult;
		}

		private static Regex HasDateTimeViewRuleRegex()
		{
			if (_dateTimeViewRulesChanged || _hasDateTimeViewRuleRegex is null)
			{
				var regexString = $@"^[ ]*({string.Join("|", DateTimeViewRules.Select(Regex.Escape))})[ ]*$";
				_hasDateTimeViewRuleRegex = new Regex(regexString, RegexOptions.Compiled);
				_dateTimeViewRulesChanged = false;
			}
			return _hasDateTimeViewRuleRegex;
		}

		private static bool HasDateTimeMacros(string data) => HasDateTimeViewRuleRegex().IsMatch(data);

		/// <summary>
		/// Format cell
		/// </summary>
		/// <param name="cell">cell to be formatted</param>
		/// <returns>Formatted text used to display as cell content</returns>
		public FormatCellResult FormatCell(Cell cell)
		{
			object data = cell.InnerData;

			bool isFormat = false;
			double number;
			DateTime value = Constants.ExcelZeroDatePoint;
			string formattedText = null;

			if (data is DateTime)
			{
				value = (DateTime)data;
				isFormat = true;
			}
			else if (CellUtility.TryGetNumberData(data, out number))
			{
				try
				{
					// Excel/Lotus 2/29/1900 bug   
					// original post: http://stackoverflow.com/questions/4538321/reading-datetime-value-from-excel-sheet
					value = DateTime.FromOADate(number);
					if (value < Constants.ExcelZeroDatePoint)
					{
						value = Constants.ExcelZeroDatePoint + value.TimeOfDay;
					}
					isFormat = true;
				}
				catch { }
			}
			else
			{
				string strdata = (data is string ? (string)data : Convert.ToString(data));

				double days = 0;
				if (double.TryParse(strdata, out days))
				{
					try
					{
						value = value.AddDays(days);
						isFormat = true;
					}
					catch { }
				}
				else
				{
					isFormat = DateTime.TryParse(strdata, out value)
								|| TryParseDateTime(CurrentCultureDateTimeFormats, strdata, out value) 
								|| TryParseDateTime(InvariantCultureDateTimeFormats, strdata, out value);
				}
			}

			if (isFormat)
			{
				if (cell.InnerStyle.HAlign == ReoGridHorAlign.General)
				{
					cell.RenderHorAlign = ReoGridRenderHorAlign.Right;
				}

				CultureInfo culture = null;

				// TODO Рассмотреть вохможность конвертирования из разных форматов
				//string pattern = System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern;
				string pattern = "dd.MM.yyyy HH:mm:ss";

				if (cell.DataFormatArgs is DateTimeFormatArgs dargs)
				{
					// fixes issue #203: pattern is ignored incorrectly
					if (!string.IsNullOrEmpty(dargs.Format))
					{
						pattern = dargs.Format;
					}

					culture = (dargs.CultureName == null
						|| string.Equals(dargs.CultureName, Thread.CurrentThread.CurrentCulture.Name))
						? Thread.CurrentThread.CurrentCulture : new CultureInfo(dargs.CultureName);
				}
				else
				{
					culture = System.Threading.Thread.CurrentThread.CurrentCulture;
					cell.DataFormatArgs = new DateTimeFormatArgs { Format = pattern, CultureName = culture.Name };
				}

				if (culture.Name.StartsWith("ja") && pattern.Contains("g"))
				{
					culture = new CultureInfo("ja-JP", true);
					culture.DateTimeFormat.Calendar = new JapaneseCalendar();
				}
				if (pattern.Contains(@"AM/PM"))
				{
					if (culture.Name.StartsWith("ru"))
					{
						pattern = pattern.Replace(@"AM/PM", string.Empty);
						pattern = pattern.Replace(@"h", @"HH");
					}
					else
					{
						pattern = pattern.Replace(@"AM/PM", @"tt");
					}
				}
				try
				{
					switch (pattern)
					{
						case "d":
							formattedText = value.Day.ToString();
							break;

						default:
							formattedText = value.ToString(pattern, culture);
							break;
					}
				}
				catch
				{
					formattedText = Convert.ToString(value);
				}
			}

			return isFormat ? new FormatCellResult(formattedText, value) 
				: (HasDateTimeMacros(Convert.ToString(data))
					? new FormatCellResult(Convert.ToString(data), data) 
					: null);
		}

		/// <summary>
		/// Represents the argument that is used during format a cell as data time.
		/// </summary>
		[Serializable]
		public struct DateTimeFormatArgs
		{
			private string format;
			/// <summary>
			/// Get or set the date time pattern. (Standard .NET datetime pattern is supported, e.g.: yyyy/MM/dd)
			/// </summary>
			public string Format { get { return format; } set { format = value; } }

			private string cultureName;
			/// <summary>
			/// Get or set the culture name that is used to format datetime according to localization settings.
			/// </summary>
			public string CultureName { get { return cultureName; } set { cultureName = value; } }

			/// <summary>
			/// Compare to another object, check whether or not two objects are same.
			/// </summary>
			/// <param name="obj">Another object to be compared.</param>
			/// <returns>True if two objects are same; Otherwise return false.</returns>
			public override bool Equals(object obj)
			{
				if (!(obj is DateTimeFormatArgs)) return false;
				DateTimeFormatArgs o = (DateTimeFormatArgs)obj;
				return format.Equals(o.format)
					&& cultureName.Equals(o.cultureName);
			}

			/// <summary>
			/// Get the hash code of this argument object.
			/// </summary>
			/// <returns>Hash code of argument object.</returns>
			public override int GetHashCode()
			{
				return format.GetHashCode() ^ cultureName.GetHashCode();
			}
		}

		/// <summary>
		/// Determines whether or not to perform a test when target cell is not set as datetime format.
		/// </summary>
		/// <returns></returns>
		public bool PerformTestFormat()
		{
			return true;
		}
	}
}
