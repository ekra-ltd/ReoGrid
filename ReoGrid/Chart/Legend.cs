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
 * Copyright (c) 2012-2023 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2023 unvell inc. All rights reserved.
 * 
 ****************************************************************************/

#if DRAWING

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Documents;

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
			// Список кандидатов на попадание в Children. Применяется для ограничеия количество элементов легенды, так
			// как в противном случае элементы легенды выходят за область отрисовки графика
			var toAddInChildrenCandidates = new List<ChartLegendItem>();

			// Класс для расчета размеров области, занимаемой легендой
			var fitCalculation = new FitCalculationHelper();
			
			Double maxSymbolWidth = 0, maxSymbolHeight = 0, maxLabelWidth = 0, maxLabelHeight = 0;

			#region Measure Sizes

			for (int index = 0; index < dataCount; index++)
			{
				var saveFitCalculation = fitCalculation;
				
				var legendItem = new ChartLegendItem(this, index);

				var symbolSize = GetSymbolSize(index);
				fitCalculation.TryUpdateMaxSymbolSize(symbolSize);
				
				if (maxSymbolWidth < symbolSize.Width) maxSymbolWidth = symbolSize.Width;
				if (maxSymbolHeight < symbolSize.Height) maxSymbolHeight = symbolSize.Height;

				legendItem.SymbolBounds = new Rectangle(new Point(0, 0), symbolSize);

				var labelSize = GetLabelSize(index);

				// should +6, don't know why
				labelSize.Width += 6;
				fitCalculation.TryUpdateMaxLabelSize(labelSize);

				if (maxLabelWidth < labelSize.Width) maxLabelWidth = labelSize.Width;
				if (maxLabelHeight < labelSize.Height) maxLabelHeight = labelSize.Height;

				legendItem.LabelBounds = new Rectangle(new Point(0, 0), labelSize);

				toAddInChildrenCandidates.Add(legendItem);

				if (!fitCalculation.IsLegendItemsFitsInClientArea(toAddInChildrenCandidates, parentClientRect, LegendPosition))
				{
					toAddInChildrenCandidates.Remove(legendItem);
					fitCalculation = saveFitCalculation;
					break;
				}
			}
			Children.AddRange(toAddInChildrenCandidates);
			
			#endregion // Measure Sizes

			#region Layout

			var itemWidth = fitCalculation.ItemWidth();
			var itemHeight = fitCalculation.ItemHeight();

			var clientRect = parentClientRect;
			Double x = 0, y = 0, right = 0, bottom = 0;

			foreach (var child in toAddInChildrenCandidates)
			{
				if (child is ChartLegendItem legendItem)
				{
					legendItem.SetSymbolLocation(0, (itemHeight - legendItem.SymbolBounds.Height) / 2);
					legendItem.SetLabelLocation(fitCalculation.MaxSymbolWidth + FitCalculationHelper.SymbolLabelSpacing, (itemHeight - legendItem.LabelBounds.Height) / 2);

					legendItem.Bounds = new Rectangle(x, y, itemWidth, itemHeight);

					if (right < legendItem.Right) right = legendItem.Right;
					if (bottom < legendItem.Bottom) bottom = legendItem.Bottom;
				}

				fitCalculation.GentNextPosition(ref x, ref y, itemWidth, itemHeight, clientRect, LegendPosition);
			}

			#endregion // Layout
			_layoutedSize = new Size(right, bottom);
		}

		#endregion

		#region Поля

		private LegendPosition _legendPosition;
		private Size _layoutedSize = Size.Zero;

		#endregion
		
		private struct FitCalculationHelper
        {
            public const double SymbolLabelSpacing = 4;
            private const double ItemSpacing = 4;

            public double MaxSymbolWidth;
            private double _maxSymbolHeight;
            private double _maxLabelWidth;
            private double _maxLabelHeight;

            public double ItemWidth()
            {
                return MaxSymbolWidth + SymbolLabelSpacing + _maxLabelWidth;
            }

            public double ItemHeight()
            {
                return Math.Max(_maxSymbolHeight, _maxLabelHeight);
            }

            private void TryUpdateMaxSymbolWidth(double value)
            {
                if (MaxSymbolWidth < value) MaxSymbolWidth = value;
            }

            private void TryUpdateMaxSymbolHeight(double value)
            {
                if (_maxSymbolHeight < value) _maxSymbolHeight = value;
            }

            public void TryUpdateMaxSymbolSize(Size size)
            {
                TryUpdateMaxSymbolWidth(size.Width);
                TryUpdateMaxSymbolHeight(size.Height);
            }

            private void TryUpdateMaxLabelWidth(double value)
            {
                if (_maxLabelWidth < value) _maxLabelWidth = value;
            }

            private void TryUpdateMaxLabelHeight(double value)
            {
                if (_maxLabelHeight < value) _maxLabelHeight = value;
            }

            public void TryUpdateMaxLabelSize(Size size)
            {
                TryUpdateMaxLabelWidth(size.Width);
                TryUpdateMaxLabelHeight(size.Height);
            }

            /// <summary>
            /// Проверка что указанные элементы легенды помещаются в отведенную для них область
            /// </summary>
            /// <param name="items"></param>
            /// <param name="clientArea"></param>
            /// <param name="legendPosition"></param>
            /// <returns></returns>
            public bool IsLegendItemsFitsInClientArea(List<ChartLegendItem> items, Rectangle clientArea,  LegendPosition legendPosition)
            {
                if (legendPosition == LegendPosition.Left || legendPosition == LegendPosition.Right)
                {
                    return (ItemHeight() + ItemSpacing) * items.Count < (clientArea.Height - 10);
                }

                if (legendPosition == LegendPosition.Bottom)
                {
                    // Здесь, вероятно, есть какой то способ линейного подсчета (без цикла), но я его еще не придумал
                    var itemHeight = ItemHeight();
                    var itemWidth = ItemWidth();
                    double x = 0, y = 0;
                    foreach (var legendItem in items)
                    {
                        GentNextPosition(ref x, ref y, itemWidth, itemHeight, clientArea, legendPosition);
                        if (x == 0)
                        {
                            // был переход на новую строку
                            if (y > clientArea.Height)
                                return false;
                        }
                        else
                        {
                            // находимся в строке с LegendItem
                            if (x + itemWidth > clientArea.Width)
                                return false;
                            if (y + itemHeight > clientArea.Height)
                                return false;
                        }
                    }
                    return true;
                }

                // Другие варианты не обрабатываются в скаде
                return true;
            }

            public void GentNextPosition(ref double x, ref double y, double itemWidth, double itemHeight, Rectangle clientArea, LegendPosition legendPosition)
            {
                x += itemWidth;

                if (legendPosition == LegendPosition.Left || legendPosition == LegendPosition.Right)
                {
                    x = 0;
                    y += itemHeight + ItemSpacing;
                }
                else
                {
                    x += ItemSpacing;

                    if (x + itemWidth > clientArea.Width)
                    {
                        x = 0;
                        y += itemHeight + ItemSpacing;
                    }
                }
            }
        }
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