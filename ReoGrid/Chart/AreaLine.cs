﻿/*****************************************************************************
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

using System.Collections.Generic;
using System.Windows;
using unvell.ReoGrid.Rendering;

namespace unvell.ReoGrid.Chart
{
	/// <summary>
	/// Represents line chart component.
	/// </summary>
	public class AreaChart : AxisChart
	{
		private AreaLineChartPlotView areaLineChartPlotView;

		/// <summary>
		/// Get plot view object of line chart component.
		/// </summary>
		public AreaLineChartPlotView AreaLineChartPlotView
		{
			get { return this.areaLineChartPlotView; }
			protected set { this.areaLineChartPlotView = value; }
		}

		/// <summary>
		/// Create line chart component instance.
		/// </summary>
		public AreaChart()
		{
			base.AddPlotViewLayer(this.areaLineChartPlotView = new AreaLineChartPlotView(this));
		}

		/// <summary>
		/// Creates and returns line chart legend instance.
		/// </summary>
		/// <returns>Instance of line chart legend.</returns>
		protected override ChartLegend CreateChartLegend(LegendType type)
		{
			return new LineChartLegend(this);
		}
	}

	public class AreaLineChartPlotView : LineChartPlotView
	{
		/// <summary>
		/// Create line chart plot view object.
		/// </summary>
		/// <param name="chart">Parent chart component instance.</param>
		public AreaLineChartPlotView(AxisChart chart)
			: base(chart)
		{
		}

		/// <summary>
		/// Render plot view region of line chart component.
		/// </summary>
		/// <param name="dc">Platform no-associated drawing context.</param>
		protected override void OnPaint(DrawingContext dc)
		{
			var axisChart = base.Chart as AxisChart;
			if (axisChart == null) return;

			var ds = Chart.DataSource;

			var g = dc.Graphics;
			var clientRect = this.ClientBounds;


#if WINFORM
			var path = new System.Drawing.Drawing2D.GraphicsPath();

			for (int r = 0; r < ds.SerialCount; r++)
			{
				var style = axisChart.DataSerialStyles[r];
				var lastPoint = new System.Drawing.PointF(axisChart.PlotColumnPoints[0], axisChart.ZeroHeight);

				for (int c = 0; c < ds.CategoryCount; c++)
				{
					var pt = axisChart.PlotDataPoints[r][c];

					System.Drawing.PointF point;

					if (pt.hasValue)
					{
						point = new System.Drawing.PointF(axisChart.PlotColumnPoints[c], axisChart.ZeroHeight - pt.value);
					}
					else
					{
						point = new System.Drawing.PointF(axisChart.PlotColumnPoints[c], axisChart.ZeroHeight);
					}

					path.AddLine(lastPoint, point);
					lastPoint = point;
				}

				var endPoint = new System.Drawing.PointF(axisChart.PlotColumnPoints[ds.CategoryCount - 1], axisChart.ZeroHeight);

				if (lastPoint != endPoint)
				{
					path.AddLine(lastPoint, endPoint);
				}

				path.CloseFigure();

				g.FillPath(style.FillColor, path);

				path.Reset();
			}

			path.Dispose();
#elif WPF


			for (int r = 0; r < ds.SerialCount; r++)
			{
				var style = axisChart.DataSerialStyles[r];

				var waZeroHeight = axisChart.ZeroHeight;
				if (waZeroHeight > clientRect.Height)
				{
					waZeroHeight = clientRect.Height;
				}
				else if (waZeroHeight < clientRect.Y)
				{
					waZeroHeight = clientRect.Y;
				}

				var seg = new System.Windows.Media.PathFigure
				{
					StartPoint = new Point(axisChart.PlotColumnPoints[0], waZeroHeight)
				};


				for (int c = 0; c < ds.CategoryCount; c++)
				{
					try
					{
						var pt = axisChart.PlotDataPoints[r][c];

						var points = new List<Point>();

						if (pt.hasValue)
						{
							if (c > 0)
							{
								var prevPt = axisChart.PlotDataPoints[r][c - 1];
								if (!prevPt.hasValue)
								{
									points.Add(new Point(axisChart.PlotColumnPoints[c], waZeroHeight));
								}
							}

							points.Add(new Point(axisChart.PlotColumnPoints[c], axisChart.ZeroHeight - pt.value));
						}
						else
						{
							if (c > 0)
							{
								var prevPt = axisChart.PlotDataPoints[r][c - 1];
								if (prevPt.hasValue)
								{
									points.Add(new Point(axisChart.PlotColumnPoints[c - 1], waZeroHeight));
								}
							}

							points.Add(new Point(axisChart.PlotColumnPoints[c], waZeroHeight));
						}

						foreach (var point in points)
						{
							seg.Segments.Add(new System.Windows.Media.LineSegment(point, true));
						}
					}
					catch
					{
						// ignored
					}
				}

				var endPoint = new System.Windows.Point(axisChart.PlotColumnPoints[ds.CategoryCount - 1], waZeroHeight);
				seg.Segments.Add(new System.Windows.Media.LineSegment(endPoint, true));

				seg.IsClosed = true;

				var path = new System.Windows.Media.PathGeometry();
				path.Figures.Add(seg);
				g.FillPath(style.LineColor, path);
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

#endif // WPF
		}

	}
}

#endif // DRAWING
