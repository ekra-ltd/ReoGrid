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

#if WINFORM || ANDROID
using RGFloat = System.Single;
#else
using RGFloat = System.Double;
#endif // WINFORM

using unvell.ReoGrid.Rendering;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid.Drawing.Shapes
{

    #region Path
    /// <summary>
    /// Represents path shape drawing object.
    /// </summary>
    public abstract class PathShape : ShapeObject
	{
		public override void OnBoundsChanged(Graphics.Rectangle oldRect)
		{
			base.OnBoundsChanged(oldRect);

			if (Width > 0 && Height > 0)
			{
				UpdatePath();
			}
		}

#if WINFORM
		protected System.Drawing.Drawing2D.GraphicsPath Path = new System.Drawing.Drawing2D.GraphicsPath();
#elif WPF
		protected System.Windows.Media.PathGeometry Path = new System.Windows.Media.PathGeometry();
#elif ANDROID
		protected Android.Graphics.Path Path = new Android.Graphics.Path();
#endif // WINFORM

		protected abstract void UpdatePath();

		/// <summary>
		/// Render path shape to graphics context.
		/// </summary>
		/// <param name="dc">Platform no-associated drawing context instance.</param>
		protected override void OnPaint(DrawingContext dc)
		{
			var g = dc.Graphics;

			if (!this.FillColor.IsTransparent)
			{
				g.FillPath(this.FillColor, this.Path);
			}

			if (!this.LineColor.IsTransparent)
			{
				g.DrawPath(this.LineColor, this.Path);
			}

			base.OnPaintText(dc);
		}

	}
	#endregion // Path

	#region Rounded Rectangle
	/// <summary>
	/// Represents a rounded rectangle shape.
	/// </summary>
	public class RoundedRectangleShape : PathShape
	{
		private RGFloat roundRate = 0.2f;

		/// <summary>
		/// Get or set the rounded corner rate relative to the minimum value between width and height. (0.0f ~ 1.0f)
		/// </summary>
		public RGFloat RoundRate
		{
			get { return roundRate; }
			set
			{
				if (this.roundRate != value)
				{
					this.roundRate = value;
					this.Invalidate();
				}
			}
		}

		protected override void UpdatePath()
		{
			RGFloat min = Math.Min(Width, Height);
			RGFloat c = roundRate * min;

#if WINFORM
			Path.Reset();

			if (c > 0)
			{
				Path.AddArc(0, 0, c, c, 180, 90);
				Path.AddArc(Width - c - 1, 0, c, c, 270, 90);
				Path.AddArc(Width - c - 1, Height - c, c, c, 0, 90);
				Path.AddArc(0, Height - c, c, c, 90, 90);
				Path.CloseAllFigures();
			}
			else
			{
				Path.AddRectangle(new System.Drawing.RectangleF(0, 0, Width, Height));
			}
#elif WPF

			Path.Clear();

			if (c > 0)
			{
				Path.AddGeometry(new System.Windows.Media.RectangleGeometry(new System.Windows.Rect(0, 0, Width, Height))
				{
					RadiusX = c,
					RadiusY = c
				});
			}

#elif ANDROID
#endif // WINFORM
		}

		protected override Rectangle TextBounds
		{
			get
			{
				RGFloat min = Math.Min(Width, Height) / 4;
				RGFloat c = roundRate * min;

				var rect = base.TextBounds;
				rect.Inflate(-c, -c);

				return rect;
			}
		}
	}
	#endregion // Rounded Rectangle

	#region Pie
	/// <summary>
	/// Represents a pie shape 
	/// </summary>
	public class PieShape : PathShape
	{
		#region Attributes
		private RGFloat startAngle = 0;

		/// <summary>
		/// Get or set the start angle of pie shape
		/// </summary>
		public virtual RGFloat StartAngle
		{
			get { return this.startAngle; }
			set
			{
				if (this.startAngle != value)
				{
					this.startAngle = value;
					this.UpdatePath();
				}
			}
		}

		private RGFloat sweepAngle = 30;

		/// <summary>
		/// Get or set the sweep angle of pie shape (Sweep from start angle)
		/// </summary>
		public virtual RGFloat SweepAngle
		{
			get { return this.sweepAngle; }
			set
			{
				if (this.sweepAngle != value)
				{
					this.sweepAngle = value;
					this.UpdatePath();
				}
			}
		}
		#endregion // Attributes

		protected override void UpdatePath()
		{
			var clientRect = this.ClientBounds;

#if WINFORM
			Path.Reset();

			if (sweepAngle > 0 && clientRect.Width > 0 && clientRect.Height > 0)
			{
				Path.AddLine(this.OriginPoint, this.OriginPoint);
				Path.AddArc(0, 0, clientRect.Width, clientRect.Height, this.startAngle - 90, this.sweepAngle);
				Path.CloseAllFigures();
			}
#elif WPF

			Path.Clear();

			if (this.sweepAngle > 0)
			{
				// 2017-12-25 исправлена отрисовка круговой диаграммы
				var pf = new System.Windows.Media.PathFigure
				{
					StartPoint = OriginPoint
				};
				double h = OriginPoint.X;

				double s = startAngle + 90; // Приведение к виду excel
				double start = (s) * Math.PI / 180.0;
				double full = (s + sweepAngle) * Math.PI / 180.0;

				start = -start; // приведение к виду excel
				full = -full;


				pf.Segments.Add(new System.Windows.Media.LineSegment(
					new System.Windows.Point(
						h * (1 - Math.Cos(start)),
						h * (1 + Math.Sin(start))),
					false));

				pf.Segments.Add(
					new System.Windows.Media.ArcSegment(new System.Windows.Point(h * (1 - Math.Cos(full)), h * (1 + Math.Sin(full))),
						new System.Windows.Size(h, h), 0, sweepAngle > 180, System.Windows.Media.SweepDirection.Clockwise, false));

				Path.Figures.Add(pf);
			}

#elif ANDROID
#endif // WINFORM
		}
	}
	#endregion // Pie

}

#endif // DRAWING