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
 * Author: Jing Lu <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2021 Jing Lu <jingwood at unvell.com>
 * Copyright (c) 2012-2016 unvell.com, all rights reserved.
 * 
 ****************************************************************************/

#if DRAWING

using System;

#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using RGFloat = System.Double;
#endif // WPF

using unvell.ReoGrid.Drawing;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;

namespace unvell.ReoGrid.Chart
{
	/// <summary>
	/// Represents chart legend view.
	/// </summary>
	public class ChartLegend : DrawingComponent
	{
		#region Конструктор

		/// <summary>
		/// Create chart legend view.
		/// </summary>
		/// <param name="chart">Instance of owner chart.</param>
		public ChartLegend(IChart chart)
		{
			Chart = chart;

			LineColor = SolidColor.Transparent;
			FillColor = SolidColor.Transparent;
			FontSize *= 0.8f;
		}

		#endregion

		#region Свойства

		/// <summary>
		/// Get or set type of legend.
		/// </summary>
		public LegendType LegendType { get; set; }



		/// <summary>
		/// Get or set the display position of legend.
		/// </summary>
		public LegendPosition LegendPosition
		{
			get => _legendPosition;
			set
			{
				if (_legendPosition != value)
				{
					_legendPosition = value;

					if (Chart is Chart)
					{
						var chart = (Chart)Chart;
						chart.DirtyLayout();
					}
				}
			}
		}

		#endregion

		#region Переопределенные методы

		/// <summary>
		/// Get measured legend view size.
		/// </summary>
		/// <returns>Measured size of legend view.</returns>
		// ReSharper disable once InheritdocConsiderUsage
		public override Size GetPreferredSize()
		{
			return _layoutedSize;
		}

		#endregion

		#region Переопределяемые методы и свойства

		protected virtual int GetLegendItemsCount()
		{
			return Chart?.DataSource?.SerialCount ?? 0;
		}

		public virtual string GetLegendLabel(int index)
		{
			return Chart.DataSource[index].Label;
		}

		/// <summary>
		/// Get the instance of owner chart.
		/// </summary>
		public virtual IChart Chart { get; protected set; }

		/// <summary>
		/// Get default symbol size of chart legend.
		/// </summary>
		/// <param name="index">Index of serial in data source.</param>
		/// <returns>Symbol size of chart legend.</returns>
		protected virtual Size GetSymbolSize(int index)
		{
			return new Size(14, 14);
		}

		/// <summary>
		/// Measure serial label size.
		/// </summary>
		/// <param name="index">Index of serial in data source.</param>
		/// <returns>Measured size for serial label.</returns>
		protected virtual Size GetLabelSize(int index)
		{
			var ds = Chart.DataSource;

			if (ds == null) return Size.Zero;

			string label = /*ds[index].Label*/GetLegendLabel(index);

			return PlatformUtility.MeasureText(null, label, FontName, FontSize, FontStyles);
		}


		/// <summary>
		/// Layout all legned items.
		/// </summary>
		public virtual void MeasureSize(Rectangle parentClientRect)
		{
			var ds = Chart.DataSource;
			if (ds == null) return;

			int dataCount = /*ds.SerialCount*/GetLegendItemsCount();

			Children.Clear();

			Double maxSymbolWidth = 0, maxSymbolHeight = 0, maxLabelWidth = 0, maxLabelHeight = 0;

			#region Measure Sizes

			for (int index = 0; index < dataCount; index++)
			{
				var legendItem = new ChartLegendItem(this, index);

				var symbolSize = GetSymbolSize(index);

				if (maxSymbolWidth < symbolSize.Width) maxSymbolWidth = symbolSize.Width;
				if (maxSymbolHeight < symbolSize.Height) maxSymbolHeight = symbolSize.Height;

				legendItem.SymbolBounds = new Rectangle(new Point(0, 0), symbolSize);

				var labelSize = GetLabelSize(index);

				// should +6, don't know why
				labelSize.Width += 6;

				if (maxLabelWidth < labelSize.Width) maxLabelWidth = labelSize.Width;
				if (maxLabelHeight < labelSize.Height) maxLabelHeight = labelSize.Height;

				legendItem.LabelBounds = new Rectangle(new Point(0, 0), labelSize);

				Children.Add(legendItem);
			}

			#endregion // Measure Sizes

			#region Layout

			const Double symbolLabelSpacing = 4;

			var itemWidth = maxSymbolWidth + symbolLabelSpacing + maxLabelWidth;
			var itemHeight = Math.Max(maxSymbolHeight, maxLabelHeight);

			var clientRect = parentClientRect;
			Double x = 0, y = 0, right = 0, bottom = 0;

			for (int index = 0; index < dataCount; index++)
			{
				var legendItem = Children[index] as ChartLegendItem;

				if (legendItem != null)
				{
					legendItem.SetSymbolLocation(0, (itemHeight - legendItem.SymbolBounds.Height) / 2);
					legendItem.SetLabelLocation(maxSymbolWidth + symbolLabelSpacing, (itemHeight - legendItem.LabelBounds.Height) / 2);

					legendItem.Bounds = new Rectangle(x, y, itemWidth, itemHeight);

					if (right < legendItem.Right) right = legendItem.Right;
					if (bottom < legendItem.Bottom) bottom = legendItem.Bottom;
				}

				x += itemWidth;

				const Double itemSpacing = 4;

				if (LegendPosition == LegendPosition.Left || LegendPosition == LegendPosition.Right)
				{
					x = 0;
					y += itemHeight + itemSpacing;
				}
				else
				{
					x += itemSpacing;

					if (x > clientRect.Width)
					{
						x = 0;
						y += itemHeight + itemSpacing;
					}
				}
			}

			#endregion // Layout

			_layoutedSize = new Size(right + 10, bottom);
		}

		#endregion

		#region Поля

		private LegendPosition _legendPosition;
		private Size _layoutedSize = Size.Zero;

		#endregion
	}

	/// <summary>
	/// Represents chart legend item.
	/// </summary>
	public class ChartLegendItem : DrawingObject
	{
		private Rectangle _symbolBounds;

		public virtual Rectangle SymbolBounds
		{
			get => _symbolBounds;
			set => _symbolBounds = value;
		}

		private Rectangle _labelBounds;

		public virtual Rectangle LabelBounds
		{
			get => _labelBounds;
			set => _labelBounds = value;
		}

		public virtual void SetSymbolLocation(Double x, Double y)
		{
			_symbolBounds.X = x;
			_symbolBounds.Y = y;
		}

		public virtual void SetLabelLocation(Double x, Double y)
		{
			_labelBounds.X = x;
			_labelBounds.Y = y;
		}

		public virtual ChartLegend ChartLegend { get; protected set; }

		public ChartLegendItem(ChartLegend chartLegend, int legendIndex)
		{
			ChartLegend = chartLegend;
			LegendIndex = legendIndex;
		}

		public virtual int LegendIndex { get; set; }

		protected override void OnPaint(DrawingContext dc)
		{
#if DEBUG
			//dc.Graphics.FillRectangle(this.ClientBounds, SolidColor.LightSteelBlue);
#endif // DEBUG

			if (_symbolBounds.Width > 0 && _symbolBounds.Height > 0)
			{
				OnPaintSymbol(dc);
			}

			if (_labelBounds.Width > 0 && _labelBounds.Height > 0)
			{
				OnPaintLabel(dc);
			}
		}

		/// <summary>
		/// Draw chart legend symbol.
		/// </summary>
		/// <param name="dc">Platform no-associated drawing context instance.</param>
		public virtual void OnPaintSymbol(DrawingContext dc)
		{
			var g = dc.Graphics;

			if (ChartLegend != null)
			{
				var legend = ChartLegend;

				if (legend.Chart != null)
				{
					var dss = legend.Chart.DataSerialStyles;

					if (dss != null)
					{
						var dsStyle = dss[LegendIndex];

						g.DrawAndFillRectangle(_symbolBounds, dsStyle.LineColor, dsStyle.FillColor);
					}
				}
			}
		}

		/// <summary>
		/// Draw chart legend label.
		/// </summary>
		/// <param name="dc">Platform no-associated drawing context instance.</param>
		public virtual void OnPaintLabel(DrawingContext dc)
		{
			if (ChartLegend != null)
			{
				var legend = ChartLegend;

				if (legend.Chart != null && legend.Chart.DataSource != null)
				{
					// var ds = legend.Chart.DataSource;
					// string itemTitle = ds[LegendIndex].Label;
					string itemTitle = ChartLegend.GetLegendLabel(LegendIndex);

					if (!string.IsNullOrEmpty(itemTitle))
					{
#if DEBUG
						//dc.Graphics.FillRectangle(this.labelBounds, SolidColor.LightCoral);
#endif // DEBUG

						dc.Graphics.DrawText(itemTitle, FontName, FontSize, ForeColor, _labelBounds,
							ReoGridHorAlign.Left, ReoGridVerAlign.Middle);
					}
				}
			}
		}
	}

	/// <summary>
	/// Legend type.
	/// </summary>
	public enum LegendType
	{
		/// <summary>
		/// Primary legend.
		/// </summary>
		PrimaryLegend,

		/// <summary>
		/// Secondary legend.
		/// </summary>
		SecondaryLegend,
	}

	/// <summary>
	/// Legend position.
	/// </summary>
	public enum LegendPosition
	{
		/// <summary>
		/// Right
		/// </summary>
		Right,

		/// <summary>
		/// Bottom
		/// </summary>
		Bottom,

		/// <summary>
		/// Left
		/// </summary>
		Left,

		/// <summary>
		/// Top
		/// </summary>
		Top,
	}
}

#endif // DRAWING