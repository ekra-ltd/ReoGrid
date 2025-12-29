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

using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;

#if WINFORM || ANDROID
using RGFloat = System.Single;
#else
using RGFloat = System.Double;
#endif // WINFORM

namespace unvell.ReoGrid.Chart
{
	/// <summary>
	/// Represents a horizontal bar chart.
	/// </summary>
	public class BarChart : ColumnChart
	{
		/// <summary>
		/// Create the instance of bar chart.
		/// </summary>
		public BarChart()
		{
			this.ShowHorizontalGuideLines = false;
			this.ShowVerticalGuideLines = true;
		}

		/// <summary>
		/// Create and return the chart plot view object.
		/// </summary>
		/// <returns></returns>
		protected override ColumnChartPlotView CreateColumnChartPlotView()
		{
			return new BarChartPlotView(this);
		}

		///// <summary>
		///// Return the default bounds for chart plot view body.
		///// </summary>
		///// <returns></returns>
		//protected override Rectangle GetPlotViewBounds(Rectangle)
		//{
		//	RGFloat width = this.Width * 0.60f;
		//	RGFloat height = this.Height - 80;

		//	return new Rectangle((this.Width - width) / 2, 50, width, height);
		//}

		/// <summary>
		/// Create and return the serial label axis info view.
		/// </summary>
		/// <param name="bodyBounds">Bounds for chart body.</param>
		/// <returns>Instance of serial label axis info view.</returns>
		protected override AxisInfoView CreatePrimaryAxisSerialLabelView(Rectangle bodyBounds)
		{
			return new AxisSerialLabelView(this, AxisTypes.Primary, AxisOrientation.Vertical)
			{
				Bounds = GetDefaultVerticalAxisInfoViewBounds(bodyBounds),
			};
		}

		/// <summary>
		/// Create and return the category label axis info view.
		/// </summary>
		/// <param name="bodyBounds">Bounds for chart body.</param>
		/// <returns>Instance of category label axis info view.</returns>
		protected override AxisInfoView CreatePrimaryAxisCategoryLabelView(Rectangle bodyBounds)
		{
			return new AxisCategoryLabelView(this, AxisTypes.Primary, AxisOrientation.Horizontal)
			{
				Bounds = GetDefaultHorizontalAxisInfoViewBounds(bodyBounds),
			};
		}

		protected override double GetHorizontalAxisLabelHeight(double diagramHeight) => 10;

		protected override void UpdateAxisLabelViewLayout(Rectangle plotRect)
		{
			const RGFloat spacing = 10;

			this.HorizontalAxisInfoView.Bounds = new Rectangle(this.ClientBounds.X, plotRect.Y - 5, 30, plotRect.Height + 10);
			this.VerticalAxisInfoView.Bounds = new Rectangle(plotRect.X, plotRect.Bottom + spacing, plotRect.Width, 10);
		}
	}

	/// <summary>
	/// Represents a bar chart plot view.
	/// </summary>
	public class BarChartPlotView : ColumnChartPlotView
	{
		/// <summary>
		/// Create instance of bar chart plot view.
		/// </summary>
		/// <param name="chart">Bar chart which holds this plot view.</param>
		public BarChartPlotView(BarChart chart)
			: base(chart)
		{
		}

		/// <summary>
		/// Render the bar chart plot view.
		/// </summary>
		/// <param name="dc">Platform no-associated drawing context instance.</param>
		protected override void OnPaint(DrawingContext dc)
		{
			var clientRect = this.ClientBounds;
			var availableHeight = clientRect.Height * 0.7;

			if (availableHeight < 20)
			{
				return;
			}

			var axisChart = base.Chart as AxisChart;
			if (axisChart == null) return;

			var ds = Chart.DataSource;

			var rows = ds.SerialCount;
			var columns = ds.CategoryCount;

			var groupColumnWidth = availableHeight / columns;
			var groupColumnSpace = ((clientRect.Height - availableHeight) / (columns + 1));
			var singleColumnHeight = groupColumnWidth / rows;

			var ai = axisChart.PrimaryAxisInfo;

			double y = groupColumnSpace;

			var g = dc.Graphics;

			for (int c = 0; c < columns; c++)
			{
				for (int r = 0; r < ds.SerialCount; r++)
				{
					var pt = axisChart.PlotDataPoints[r][c];

					if (pt.hasValue)
					{
						var scaleY = Bounds.Width / Bounds.Height;
						var style = axisChart.DataSerialStyles[r];
						
						try
						{
							if (pt.value > 0)
							{
								if (-axisChart.ZeroWidth < clientRect.X)
								{
									var waZeroWitdh = clientRect.X;
									var waWidth = pt.value * scaleY + (waZeroWitdh - axisChart.ZeroWidth);
									if (waWidth >= 0)
									{
										g.DrawAndFillRectangle(new Rectangle(
											waZeroWitdh, (RGFloat) y,
											waWidth, ReduceBarHeight(singleColumnHeight)), style.LineColor, style.FillColor);
									}
								}
								else
								{
									g.DrawAndFillRectangle(new Rectangle(
										-axisChart.ZeroWidth, (RGFloat) y,
										pt.value * scaleY, ReduceBarHeight(singleColumnHeight)), style.LineColor, style.FillColor);
								}
							}
							else
							{
								var xPos = -(axisChart.ZeroWidth - pt.value * scaleY);

								if (-axisChart.ZeroWidth > clientRect.Width)
								{
									var waWidth = clientRect.Width - xPos;
									if (waWidth >= 0)
									{
										g.DrawAndFillRectangle(new Rectangle(
											-(axisChart.ZeroWidth - pt.value * scaleY), (RGFloat) y,
											waWidth, ReduceBarHeight(singleColumnHeight)), style.LineColor, style.FillColor);
									}
								}
								else
								{
									g.DrawAndFillRectangle(new Rectangle(
										xPos, (RGFloat) y,
										(-pt.value) * scaleY, ReduceBarHeight(singleColumnHeight)), style.LineColor, style.FillColor);
								}
							}
						}
						catch
						{
							// ignored
						}
					}

					y += singleColumnHeight;
				}

				y += groupColumnSpace;
			}
			if (axisChart.VerticalAxisInfoView.ReverseOrderOfCategories &&
			    !axisChart.HorizontalAxisInfoView.ReverseOrderOfCategories)
			{
				dc.Graphics.ReflectionXYTransform(axisChart.VerticalAxisInfoView.Right, axisChart.HorizontalAxisInfoView.Bottom);
				return;
			}
			if(axisChart.VerticalAxisInfoView.ReverseOrderOfCategories)
				dc.Graphics.ReflectionXTransform(axisChart.VerticalAxisInfoView.Right);
			if(!axisChart.HorizontalAxisInfoView.ReverseOrderOfCategories)
				dc.Graphics.ReflectionYTransform(axisChart.HorizontalAxisInfoView.Bottom);
		}
	}
}

#endif // DRAWING