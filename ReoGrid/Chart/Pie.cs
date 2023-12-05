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

/*
 * Eсть два варианта отображения круговой диаграммы:
 *     Эталонный, принятый в excel (отображаемые в диаграмме значения берутся из первого ряда данных)
 *     реализованный в reogrid (отображаемые в диаграмме значения берутся из первого значения каждого ряда)
 * К сожалению, если реализовавывать правильный вариант - то придется дорабатываться студию и вставлять два костыля (один 
 * в студию, другой в скада). Так как изначально настройка была реализована по варианту reogrid. Кастыли в студии 
 * выглядят ужасно, поэтому лучше вставить один кастыль сюда в код экспорта в excel
 *     
 *     WORKAROUND_PIE_CHART - указывает на то что используется кастыль в reogrid,
 *     иначе используется вариант excel
 */

#define WORKAROUND_PIE_CHART

using System;
using System.Collections.Generic;
using System.Linq;

#if WINFORM || ANDROID
using RGFloat = System.Single;
#else
using RGFloat = System.Double;
#endif // WINFORM

using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Drawing.Shapes;
using unvell.ReoGrid.Drawing;

namespace unvell.ReoGrid.Chart
{
	/// <summary>
	/// Repersents pie chart component.
	/// </summary>
	public class PieChart : Chart
	{
		internal virtual PieChartDataInfo DataInfo { get; private set; }
		internal virtual List<RGFloat> PlotPieAngles { get; private set; }
		internal virtual PiePlotView PiePlotView { get; private set; }

		internal bool UseReogridWorkaround
		{
			get
			{
#if WORKAROUND_PIE_CHART
				return true;
#else
                return false;
#endif
			}
		}
		/// <summary>
		/// Creates pie chart instance.
		/// </summary>
		public PieChart()
		{
			this.DataInfo = new PieChartDataInfo();
			this.PlotPieAngles = new List<RGFloat>();
			
			this.AddPlotViewLayer(this.PiePlotView = CreatePlotViewInstance());
		}

		#region Legend
		protected override ChartLegend CreateChartLegend(LegendType type)
		{
			var chartLegend = UseReogridWorkaround ? new PieChartLegend2(this) as ChartLegend : new PieChartLegend(this);

			if (type == LegendType.PrimaryLegend)
			{
				chartLegend.LegendPosition = LegendPosition.Bottom;
			}

			return chartLegend;
		}
		#endregion // Legend

		#region Plot view instance
		/// <summary>
		/// Creates and returns pie plot view.
		/// </summary>
		/// <returns></returns>
		protected virtual PiePlotView CreatePlotViewInstance()
		{
			return new PiePlotView(this);
		}
		#endregion // Plot view instance

		#region Layout
		protected override Rectangle GetPlotViewBounds(Rectangle bodyBounds)
		{
			RGFloat minSize = Math.Min(bodyBounds.Width, bodyBounds.Height);

			return new Rectangle(bodyBounds.X + (bodyBounds.Width - minSize) / 2, 
				bodyBounds.Y + (bodyBounds.Height - minSize) / 2,
				minSize, minSize);
		}
		#endregion // Layout

		#region Data Serials
		//public override int GetSerialCount()
		//{
		//	return this.DataSource.CategoryCount;
		//}
		//public override string GetSerialName(int index)
		//{
		//	return this.DataSource == null ? string.Empty : this.DataSource.GetCategoryName(index);
		//}
		#endregion // Data Serials

		#region Update Draw Points
		//protected override int GetSerialStyleCount()
		//{
		//	var ds = this.DataSource;
		//	return ds == null ? 0 : ds.CategoryCount;
		//}

		/// <summary>
		/// Update data serial information.
		/// </summary>
		protected override void UpdatePlotData()
		{
			if (DataSource == null) return;
			DataInfo.Total = CalculateTotal();
			UpdatePlotPoints();
		}

		private double CalculateTotal() =>
			EnumeratePieValues().Sum();
		
		/// <summary>
        /// Update plot calculation points.
        /// </summary>
        protected virtual void UpdatePlotPoints()
        {
            if (DataSource != null && DataSource.SerialCount > 0)
            {
                RGFloat scale = (RGFloat)(360.0 / DataInfo.Total);
                var i = 0;
                foreach (var val in EnumeratePieValues())
                {
                    var angle = (RGFloat)(val * scale);

                    if (i >= PlotPieAngles.Count)
                    {
                        PlotPieAngles.Add(angle);
                    }
                    else
                    {
                        PlotPieAngles[i] = angle;
                    }
                    i++;
                }
            }
            else
            {
                PlotPieAngles.Clear();
            }

            PiePlotView?.Invalidate();
        }

        private IEnumerable<RGFloat> EnumeratePieValues()
        {
            if (UseReogridWorkaround)
            {
                for (var index = 0; index < DataSource.SerialCount; index++)
                {
                    if (DataSource[index][0] is RGFloat data)
                    {
                        yield return data;
                    }
                    else
                    {
                        yield return (RGFloat) 0;
                    }
                }
            }
            else
            {
                if (DataSource.SerialCount > 0)
                {
                    for (var index = 0; index < DataSource[0].Count; index++)
                    {
                        if (DataSource[0][index] is RGFloat data)
                        {
                            yield return data;
                        }
                        else
                        {
                            yield return (RGFloat)0;
                        }
                    }
                }
            }
        }
		#endregion // Update Draw Points
		
		#region Переопределенные методы

		protected override void ResetDataSerialStyles()
		{
			if (UseReogridWorkaround)
			{
				base.ResetDataSerialStyles();
			}
			else
			{
				var ds = DataSource;
				if (ds == null) return;

				if (ds.SerialCount <= 0 || ds[0].Count <= 0)
				{
					serialStyles.Clear();
				}
				int dataSerialCount = ds[0].Count;
				while (serialStyles.Count < dataSerialCount)
				{
					serialStyles.Add(new DataSerialStyle(this)
					{
						FillColor = ChartUtility.GetDefaultDataSerialFillColor(serialStyles.Count),
						LineColor = ChartUtility.GetDefaultDataSerialFillColor(serialStyles.Count),
						LineWidth = 2f,
					});
				}
			}
		}

		#endregion
	}
	
	/// <summary>
	/// Represents pie chart data information.
	/// </summary>
	public class PieChartDataInfo
	{
		public double Total { get; set; }
	}

	/// <summary>
	/// Represents pie plot view.
	/// </summary>
	public class PiePlotView : ChartPlotView
	{
		/// <summary>
		/// Create plot view object of pie 2d chart.
		/// </summary>
		/// <param name="chart">Pie chart instance.</param>
		public PiePlotView(Chart chart)
			: base(chart)
		{
			this.Chart.DataSourceChanged += Chart_DataSourceChanged;
			this.Chart.ChartDataChanged += Chart_DataSourceChanged;
		}

		~PiePlotView()
		{
			this.Chart.DataSourceChanged -= Chart_DataSourceChanged;
			this.Chart.ChartDataChanged -= Chart_DataSourceChanged;
		}

		void Chart_DataSourceChanged(object sender, EventArgs e)
		{
			this.UpdatePieShapes();
		}

		protected List<PieShape> PlotPieShapes = new List<PieShape>();

		protected virtual void UpdatePieShapes()
		{
			if (Chart is PieChart pieChart)
			{
				var ds = Chart.DataSource;
				if (ds != null && ds.SerialCount > 0)
				{
					var dataCount = 0;
					if (pieChart.UseReogridWorkaround)
					{
						dataCount = ds.SerialCount;
					}
					else
					{
						dataCount = ds[0].Count;
					}
					RGFloat currentAngle = 0;

					for (var i = 0; i < dataCount && i < pieChart.PlotPieAngles.Count; i++)
					{
						RGFloat angle = pieChart.PlotPieAngles[i];

						if (i >= PlotPieShapes.Count)
						{
							PlotPieShapes.Add(CreatePieShape(ClientBounds));
						}

						var pie = PlotPieShapes[i];
						pie.StartAngle = currentAngle;
						pie.SweepAngle = angle;
						pie.FillColor = pieChart.DataSerialStyles[i].FillColor;

						currentAngle += angle;
					}
				}
			}
		}

		protected virtual PieShape CreatePieShape(Rectangle bounds)
		{
			return new PieShape()
			{
				Bounds = bounds,
				LineColor = SolidColor.Transparent,
			};
		}

		/// <summary>
		/// Render pie 2d plot view.
		/// </summary>
		/// <param name="dc">Platform no-associated drawing context instance.</param>
		protected override void OnPaint(Rendering.DrawingContext dc)
		{
			base.OnPaint(dc);

			foreach (var pieShape in this.PlotPieShapes)
			{
				pieShape.Draw(dc);
			}
		}
	}

	/// <summary>
	/// Repersents pie 2D chart component.
	/// </summary>
	public class Pie2DChart : PieChart
	{
	}

	/// <summary>
	/// Represents pie 2D plot view.
	/// </summary>
	public class Pie2DPlotView : PiePlotView
	{
		public Pie2DPlotView(Pie2DChart pieChart)
			: base(pieChart)
		{
		}
	}

	/// <summary>
	/// Repersents doughnut chart component.
	/// </summary>
	public class DoughnutChart : PieChart
	{
		/// <summary>
		/// Creates and returns doughnut plot view.
		/// </summary>
		protected override PiePlotView CreatePlotViewInstance()
		{
			return new DoughnutPlotView(this);
		}
	}

	/// <summary>
	/// Represents doughnut plot view.
	/// </summary>
	public class DoughnutPlotView : PiePlotView
	{
		public DoughnutPlotView(DoughnutChart chart)
			: base(chart)
		{
		}

		protected override PieShape CreatePieShape(Rectangle bounds)
		{
			return new Drawing.Shapes.SmartShapes.BlockArcShape
			{
				Bounds = bounds,
				LineColor = SolidColor.White,
			};
		}
	}
	
	public class PieChartLegend: ChartLegend
	{
		public PieChartLegend(IChart chart) : base(chart)
		{
		}

		#region Переопределенные методы

		protected override int GetLegendItemsCount()
		{
			var serialCount = Chart?.DataSource?.SerialCount;
			if (serialCount.HasValue && serialCount.Value >= 0)
			{
				return Chart.DataSource[0].Count;
			}
			return 0;
		}

		public override string GetLegendLabel(int index)
		{
			if (index >= 0 && index < Chart.DataSource.CategoryCount)
			{
				return Chart.DataSource.GetCategoryName(index) ?? string.Empty;
			}
			return string.Empty;
		}


		#endregion
	}
}

#endif // DRAWING