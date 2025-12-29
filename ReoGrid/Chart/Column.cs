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

#if DRAWING

using System;
using System.Windows.Media;

#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using RGFloat = System.Double;
#endif // WPF

using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;
using DrawingContext = unvell.ReoGrid.Rendering.DrawingContext;

namespace unvell.ReoGrid.Chart
{
	/// <summary>
	/// Represents column chart component.
	/// </summary>
	public class ColumnChart : AxisChart
	{
		private ColumnChartPlotView chartPlotView;

		/// <summary>
		/// Get column chart plot view object.
		/// </summary>
		public ColumnChartPlotView ColumnChartPlotView
		{
			get { return this.chartPlotView; }
			protected set { this.chartPlotView = value; }
		}

		/// <summary>
		/// Create column chart instance.
		/// </summary>
		public ColumnChart()
		{
			base.AddPlotViewLayer(this.chartPlotView = CreateColumnChartPlotView());
		}

		/// <summary>
		/// Create and return the main plot view for this chart. 
		/// Derived classes specify their own plot view objects by overwriting this method.
		/// </summary>
		/// <returns>Plot view object for column-based charts.</returns>
		protected virtual ColumnChartPlotView CreateColumnChartPlotView()
		{
			return new ColumnChartPlotView(this);
	}

		/// <summary>
		/// Creates and returns column line chart legend instance.
		/// </summary>
		/// <returns>Instance of column line chart legend.</returns>
		protected override ChartLegend CreateChartLegend(LegendType type)
		{
			return new ColumnLineChartLegend(this);
		}

		protected new virtual void UpdateAxisInfo(AxisDataInfo ai, double minData, double maxData)
		{
			var clientRect = this.PlotViewContainer;

			double range = maxData - minData;

			ai.Levels = (int)Math.Ceiling(clientRect.Height / 30f);

			// when clientRect is zero, nothing to do
			if (double.IsNaN(ai.Levels))
			{
				return;
			}

			if (minData == maxData)
			{
				if (maxData == 0)
					maxData = ai.Levels;
				else
					minData = 0;
			}

			int scaler;
			double stride = ChartUtility.CalcLevelStride(minData, maxData, ai.Levels, out scaler);
			ai.Scaler = scaler;

			double m;

			if (!ai.AutoMinimum)
			{
				if (this.AxisOriginToZero(minData, maxData, range))
				{
					ai.Minimum = 0;
				}
				else
				{
					m = minData % stride;
					if (m == 0)
					{
						if (minData == 0)
						{
							ai.Minimum = minData;
						}
						else
						{
							ai.Minimum = minData - stride;
						}
					}
					else
					{
						if (minData < 0)
						{
							ai.Minimum = minData - stride - m;
						}
						else
						{
							ai.Minimum = minData - m;
						}
					}
				}
			}

			if (!ai.AutoMaximum)
			{
				m = maxData % stride;
				if (m == 0)
				{
					ai.Maximum = maxData + stride;
				}
				else
				{
					ai.Maximum = maxData - m + stride;
				}
			}

			ai.Levels = (int)Math.Round((ai.Maximum - ai.Minimum) / stride);

			ai.LargeStride = stride;
		}
	}

	/// <summary>
	/// Represents column line chart legend.
	/// </summary>
	public class ColumnLineChartLegend : ChartLegend
	{
		/// <summary>
		/// Create column line chart legend.
		/// </summary>
		/// <param name="chart">Parent chart component.</param>
		internal ColumnLineChartLegend(IChart chart)
			: base(chart)
		{
		}

		/// <summary>
		/// Get default symbol size of chart legend.
		/// </summary>
		/// <returns>Symbol size of chart legend.</returns>
		protected override Size GetSymbolSize(int index)
		{
			return new Size(12, 12);
		}
	}

	/// <summary>
	/// Represents plot view object of column chart component.
	/// </summary>
	public class ColumnChartPlotView : ChartPlotView
	{
		/// <summary>
		/// Create column chart plot view object.
		/// </summary>
		/// <param name="chart">Owner chart instance.</param>
		public ColumnChartPlotView(AxisChart chart)
			: base(chart)
		{
		}

		/// <summary>
		/// Render the column chart plot view.
		/// </summary>
		/// <param name="dc">Platform no-associated drawing context instance.</param>
		protected override void OnPaint(DrawingContext dc)
		{
			base.OnPaint(dc);

			//var bottomAxis = axisChart.BottomAxisInfo;
			var clientRect = this.ClientBounds;
			var availableWidth = clientRect.Width * 0.7;

			if (availableWidth < 20)
			{
				return;
			}

			var axisChart = base.Chart as AxisChart;
			if (axisChart == null) return;

			var ds = Chart.DataSource;

			var rows = ds.SerialCount;
			var columns = ds.CategoryCount;

			var roundColumnWidth = availableWidth / columns;
			var roundColumnSpace = ((clientRect.Width - availableWidth) / (columns + 1));
			var singleColumnWidth = roundColumnWidth / rows;

			var ai = axisChart.PrimaryAxisInfo;

			double x = roundColumnSpace;

			var g = dc.Graphics;

			for (int c = 0; c < columns; c++)
			{
				for (int r = 0; r < ds.SerialCount; r++)
				{
					var pt = axisChart.PlotDataPoints[r][c];

					if (pt.hasValue)
					{
						var style = axisChart.DataSerialStyles[r];

						try
						{
							if (pt.value > 0)
							{
								// Определить что график выходит за границу отрисовки (больше высоты)
								if (axisChart.ZeroHeight > clientRect.Height)
								{
									var waHeight = pt.value + (clientRect.Height - axisChart.ZeroHeight);
									if (waHeight >= 0)
									{
										g.DrawAndFillRectangle(new Rectangle(
											(RGFloat)x, axisChart.ZeroHeight - pt.value,
											ReduceBarWidth(singleColumnWidth), waHeight), style.LineColor, style.FillColor);
									}
								}
								else
								{
									g.DrawAndFillRectangle(new Rectangle(
										(RGFloat)x, axisChart.ZeroHeight - pt.value,
										ReduceBarWidth(singleColumnWidth), pt.value), style.LineColor, style.FillColor);
								}
							}
							else
							{
								// Определить что график выходит за границу отрисовки (меньше начальной точки)
								if (axisChart.ZeroHeight < clientRect.Y)
								{
									var waZeroHeight = clientRect.Y;
									var waHeight = -pt.value - (waZeroHeight - axisChart.ZeroHeight);

									if (waHeight >= 0)
									{
										g.DrawAndFillRectangle(new Rectangle(
											(RGFloat)x, waZeroHeight,
											ReduceBarWidth(singleColumnWidth), waHeight), style.LineColor, style.FillColor);
									}
								}
								else
								{

									g.DrawAndFillRectangle(new Rectangle(
										(RGFloat)x, axisChart.ZeroHeight,
										ReduceBarWidth(singleColumnWidth), -pt.value), style.LineColor, style.FillColor);
								}
							}
						}
						catch
						{
							// ignored
						}
					}
					x += singleColumnWidth;
				}

				x += roundColumnSpace;
			}
			if (axisChart.VerticalAxisInfoView.ReverseOrderOfCategories &&
			    axisChart.HorizontalAxisInfoView.ReverseOrderOfCategories)
			{
				dc.Graphics.ReflectionXYTransform(
					axisChart.HorizontalAxisInfoView.Right - axisChart.VerticalAxisInfoView.Width,
					Chart.Height
				);
				return;
			}
			if(axisChart.VerticalAxisInfoView.ReverseOrderOfCategories)
				dc.Graphics.ReflectionYTransform(Chart.Height);
			if(axisChart.HorizontalAxisInfoView.ReverseOrderOfCategories)
				dc.Graphics.ReflectionXTransform(axisChart.HorizontalAxisInfoView.Right - axisChart.VerticalAxisInfoView.Width);
		}
		
		protected static RGFloat ReduceBarHeight(RGFloat fullHeight) => ReduceValue(fullHeight);

		private static RGFloat ReduceBarWidth(RGFloat fullWidth) => ReduceValue(fullWidth);

		private static RGFloat ReduceValue(RGFloat value)
		{
			if (value > 2)
				return value - 1;
			if (value > 0.1)
				return 0.95 * value;
			return value;
		}
	}
}

#endif // DRAWING