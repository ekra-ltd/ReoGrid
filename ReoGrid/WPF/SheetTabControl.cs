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

#if WPF

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using unvell.ReoGrid.Main;
using unvell.ReoGrid.Views;

namespace unvell.ReoGrid.WPF
{
	internal class SheetTabControl : Grid, ISheetTabControl
	{
		internal Grid canvas = new Grid()
		{
			Width = 0,
			VerticalAlignment = VerticalAlignment.Top,
			HorizontalAlignment = HorizontalAlignment.Left
		};

		private Image newSheetImage;

        public SheetTabControl()
        {
			this.Background = SystemColors.ControlBrush;
			this.BorderColor = Colors.DeepSkyBlue;

			this.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
			this.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
			this.ColumnDefinitions.Add(new ColumnDefinition { });
			this.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
			this.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(5) });

            Border pleft = new ArrowBorder(this)
            {
                Child = new Polygon()
                {
                    Points = new PointCollection(
                        new Point[]{
                            new Point(6, 0),
                            new Point(0, 5),
                            new Point(6, 10),
                        }),
                    Fill = SystemColors.ControlDarkDarkBrush,
					HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
					VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new Thickness(4, 0, 0, 0),
                },
                Background = SystemColors.ControlBrush,
            };

            Border pright = new ArrowBorder(this)
            {
                Child = new Polygon()
                {
                    Points = new PointCollection(
                        new Point[] {
                            new Point(0, 0),
                            new Point(6, 5),
                            new Point(0, 10),
                        }),
                    Fill = SystemColors.ControlDarkDarkBrush,
					HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
					VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 4, 0),
                },
                Background = SystemColors.ControlBrush,
            };

			this.canvas.RenderTransform = new TranslateTransform(0, 0);

			this.Children.Add(this.canvas);
			Grid.SetColumn(this.canvas, 2);

			this.Children.Add(pleft);
			Grid.SetColumn(pleft, 0);

			this.Children.Add(pright);
			Grid.SetColumn(pright, 1);

            BitmapImage imageSource = new BitmapImage();
            BitmapImage imageHoverSource = new BitmapImage();

            imageSource.BeginInit();
			imageSource.StreamSource = new System.IO.MemoryStream(unvell.ReoGrid.Properties.Resources.NewBuildDefinition_8952_inactive_png);
            imageSource.EndInit();

            imageHoverSource.BeginInit();
			imageHoverSource.StreamSource = new System.IO.MemoryStream(unvell.ReoGrid.Properties.Resources.NewBuildDefinition_8952_png);
            imageHoverSource.EndInit();

            newSheetImage = new Image()
            {
                Source = imageSource,
				HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
				VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new Thickness(2),
				Cursor = System.Windows.Input.Cursors.Hand,
            };

            newSheetImage.MouseEnter += (s, e) => newSheetImage.Source = imageHoverSource;
            newSheetImage.MouseLeave += (s, e) => newSheetImage.Source = imageSource;
            newSheetImage.MouseDown += (s, e) =>
            {
				if (this.NewSheetClick != null)
                {
					this.NewSheetClick(this, null);
                }
            };

			this.Children.Add(newSheetImage);
			Grid.SetColumn(newSheetImage, 3);

            Border rightThumb = new Border
            {
                Child = new RightThumb(this),
				Cursor = System.Windows.Input.Cursors.SizeWE,
                Background = SystemColors.ControlBrush,
                Margin = new Thickness(0, 1, 0, 0),
				HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            };
			this.Children.Add(rightThumb);
			Grid.SetColumn(rightThumb, 4);

			this.scrollTimer = new System.Windows.Threading.DispatcherTimer()
            {
                Interval = new TimeSpan(0, 0, 0, 0, 10),
            };

            scrollTimer.Tick += (s, e) =>
            {
				var tt = this.canvas.Margin.Left;

				if (this.scrollLeftDown)
                {
                    if (tt < 0)
                    {
                        tt += 5;
                        if (tt > 0) tt = 0;
                    }
                }
				else if (this.scrollRightDown)
                {
					double max = this.ColumnDefinitions[2].ActualWidth - this.canvas.Width;

                    if (tt > max)
                    {
                        tt -= 5;
                        if (tt < max) tt = max;
                    }
                }

				if (this.canvas.Margin.Left != tt)
                {
					this.canvas.Margin = new Thickness(tt, 0, 0, 0);
                }
            };

            pleft.MouseDown += (s, e) =>
            {
				this.scrollRightDown = false;
				if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
                {
					this.scrollLeftDown = true;
					this.scrollTimer.IsEnabled = true;
                }
				else if (e.RightButton == System.Windows.Input.MouseButtonState.Pressed)
                {
					if (this.SheetListClick != null)
                    {
						this.SheetListClick(this, null);
                    }
                }
            };
            pleft.MouseUp += (s, e) =>
            {
				this.scrollTimer.IsEnabled = false;
				this.scrollLeftDown = false;
            };

            pright.MouseDown += (s, e) =>
            {
				this.scrollLeftDown = false;
				if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
                {
					this.scrollRightDown = true;
					this.scrollTimer.IsEnabled = true;
                }
				else if (e.RightButton == System.Windows.Input.MouseButtonState.Pressed)
                {
					if (this.SheetListClick != null)
                    {
						this.SheetListClick(this, null);
                    }
                }
            };
            pright.MouseUp += (s, e) =>
            {
				this.scrollTimer.IsEnabled = false;
				this.scrollRightDown = false;
            };

            rightThumb.MouseDown += (s, e) =>
            {
				this.splitterMoving = true;
                rightThumb.CaptureMouse();
            };
            rightThumb.MouseMove += (s, e) =>
            {
				if (this.splitterMoving)
                {
					if (this.SplitterMoving != null)
                    {
						this.SplitterMoving(this, null);
                    }
                }
            };
            rightThumb.MouseUp += (s, e) =>
            {
				this.splitterMoving = false;
                rightThumb.ReleaseMouseCapture();
            };
            UpdateTabsState();
		}

		private bool splitterMoving = false;
		private bool scrollLeftDown = false, scrollRightDown = false;

		private System.Windows.Threading.DispatcherTimer scrollTimer;

		protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
		{
			base.OnRenderSizeChanged(sizeInfo);

			this.Clip = new RectangleGeometry(new Rect(0, 0, this.RenderSize.Width, this.RenderSize.Height));

			this.canvas.Height = this.Height - 2;
        }

		protected override void OnInitialized(EventArgs e)
		{
			base.OnInitialized(e);
		}

		#region Dependency Properties

		public static readonly DependencyProperty SelectedBackColorProperty =
			DependencyProperty.Register("SelectedBackColor", typeof(Color), typeof(SheetTabControl));

		public Color SelectedBackColor
		{
			get { return (Color)GetValue(SelectedBackColorProperty); }
			set { SetValue(SelectedBackColorProperty, value); }
		}

		public static readonly DependencyProperty SelectedTextColorProperty =
			DependencyProperty.Register("SelectedTextColor", typeof(Color), typeof(SheetTabControl));

		public Color SelectedTextColor
		{
			get { return (Color)GetValue(SelectedTextColorProperty); }
			set { SetValue(SelectedTextColorProperty, value); }
		}

		public static readonly DependencyProperty BorderColorProperty =
			DependencyProperty.Register("BorderColor", typeof(Color), typeof(SheetTabControl));

		public Color BorderColor
	    {
			get { return (Color)GetValue(BorderColorProperty); }
			set { SetValue(BorderColorProperty, value); }
		}

		public static readonly DependencyProperty SelectedIndexProperty =
			DependencyProperty.Register("SelectedIndex", typeof(int), typeof(SheetTabControl));

		public int SelectedIndex
		{
			get { return (int)GetValue(SelectedIndexProperty); }

			set
			{
				var tabContainer = this.canvas;

				var currentIndex = this.SelectedIndex;

				if (currentIndex >= 0 && currentIndex < tabContainer.Children.Count)
				{
					var tab = tabContainer.Children[currentIndex] as SheetTabItem;
					if (tab != null)
					{
						tab.IsSelected = false;
	    }
				}

				SetValue(SelectedIndexProperty, value);
				currentIndex = value;

				if (currentIndex >= 0 && currentIndex < tabContainer.Children.Count)
				{
					var tab = tabContainer.Children[currentIndex] as SheetTabItem;
					if (tab != null)
					{
						tab.IsSelected = true;
					}
				}

				if (this.SelectedIndexChanged != null)
				{
					this.SelectedIndexChanged(this, null);
				}
			}
		}

		public bool NewButtonVisible
		{
			get { return this.newSheetImage.Visibility == Visibility.Visible; }
			set { this.newSheetImage.Visibility = value ? Visibility.Visible : Visibility.Hidden; }
		}
        
		#endregion // Dependency Properties

		/// <summary>
		/// Determine whether or not allow to move tab by dragging mouse
		/// </summary>
		public bool AllowDragToMove { get; set; }

		#region Tab Management
		public void AddTab(string title)
		{
			int index = this.canvas.Children.Count;
			InsertTab(index, title);
		}

	    public void InsertTab(int index, string title)
	    {
	        try
	        {
	            var tab = new SheetTabItem(this, title)
	            {
					Height = this.canvas.Height,
	            };
                tab.TabRenaming += OnTabRenaming;
                tab.TabRename += OnTabRename;
	            tab.TabRenamed += OnTabRenamed;

				this.canvas.Width += tab.Width + 1;
				this.canvas.ColumnDefinitions.Add(new ColumnDefinition {Width = new GridLength(tab.Width + 1)});

				this.canvas.Children.Add(tab);

				Grid.SetColumn(tab, index);

				tab.MouseDown += Tab_MouseDown;

				if (this.canvas.Children.Count == 1)
	            {
					tab.IsSelected = true;
				}
			}
			finally
			{
				UpdateTabsState();
			}
		}

		private void Tab_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
		{
			var index = this.canvas.Children.IndexOf((UIElement)sender);

	                var arg = new SheetTabMouseEventArgs()
	                {
	                    Handled = false,
	                    Location = e.GetPosition(this),
				Index = index,
	                    MouseButtons = WPFUtility.ConvertToUIMouseButtons(e),
	                };

			if (this.TabMouseDown != null)
	                {
				this.TabMouseDown(this, arg);
	                }

			if (!arg.Handled)
	            {
				this.SelectedIndex = index;
	            }
	        }

	    public void RemoveTab(int index)
        {
            try
            {
				var tab = (SheetTabItem) this.canvas.Children[index];

				this.canvas.Children.RemoveAt(index);
				this.canvas.ColumnDefinitions.RemoveAt(index);

				for (int i = index; i < this.canvas.Children.Count; i++)
                {
					Grid.SetColumn(this.canvas.Children[i], i);
					this.canvas.Children[i].InvalidateVisual(); // TODO merged
                        }

				this.canvas.Width -= tab.Width;
                    }
            finally
            {
                UpdateTabsState();
            }
        }

        /// <summary>
        /// Запрос на удаление листа
        /// </summary>
        /// <param name="item">удаляемый лист</param>
        /// <returns>true, если выполнено удаление, иначе - false</returns>
        internal bool SheetTabRemoveRequest(SheetTabItem item)
        {
            if (item == null) return false;
            var tabContainer = canvas;
            var index = tabContainer.Children.IndexOf(item);
            if (index < 0) return false;
            if (!item.IsCanRemove)
                return false;

            var e = new SheetTabRemovingEventArgs { Index = index, Cancel = false };
            SheetTabRemoving?.Invoke(this, e);
            if (!e.Cancel)
            {
                SheetTabRemove?.Invoke(this, new SheetTabRemoveEventArgs { Index = index });
                return true;
            }
            return false;
        }

        public void UpdateTab(int index, string title, Color backColor, Color textColor)
    {
			SheetTabItem item = this.canvas.Children[index] as SheetTabItem;
			if (item != null)
      {
        double oldWidth = item.Width;
        item.ChangeTitle(title);
        double newWidth = item.Width;

				this.canvas.ColumnDefinitions[index].Width = new GridLength(item.Width+1);

        item.BackColor = backColor;
        item.TextColor = textColor;

        //this.canvas.Width = this.canvas.Width - oldWidth + newWidth;
        
            double width = 1;
            foreach (UIElement uiElement in canvas.Children )
            {
                if (uiElement is SheetTabItem tab)
                    width += tab.Width;
            }
            canvas.Width = width;
        
        for (int i = index; i < this.canvas.Children.Count; i++)
        {
          Grid.SetColumn(this.canvas.Children[i], i);
        }
      }
    }

    public void ClearTabs()
        {
            try
            {
				this.canvas.Children.Clear();
				this.canvas.ColumnDefinitions.Clear();
				this.canvas.Width = 0;
            }
            finally
            {
                UpdateTabsState();
            }
        }

		public int TabCount { get { return this.canvas.Children.Count; } }
        
		/// <summary>
		/// Функция обновляет состояние Tab-ов (разрешает/запрещает их удаление)
		/// </summary>
		private void UpdateTabsState()
	    {
			var count = TabCount;
			foreach (var uiElement in canvas.Children)
	        {
				if (uiElement is SheetTabItem tab)
					tab.IsCanRemove = count > 1;
	        }
	    }

        #endregion // Tab Management

        #region Paint

        private GuidelineSet gls = new GuidelineSet();

	    protected override void OnRender(DrawingContext dc)
	    {
	        base.OnRender(dc);

	        var g = dc;

	        gls.GuidelinesX.Clear();
	        gls.GuidelinesY.Clear();

	        gls.GuidelinesX.Add(0.5);
	        gls.GuidelinesX.Add(RenderSize.Width + 0.5);
	        gls.GuidelinesY.Add(0.5);
	        gls.GuidelinesY.Add(RenderSize.Height + 0.5);

	        g.PushGuidelineSet(gls);

			var p = new Pen(new SolidColorBrush(this.BorderColor), 1);

			g.DrawLine(p, new Point(0, 0), new Point(this.RenderSize.Width, 0));

	        g.Pop();

	    }
        #endregion // Paint

		#region Переименование вкладки
        
		public bool IsInEditMode
	    {
			get;
			private set;
	        }

		public Action<int> RenameSheetTabItemCallback { get; set; }

		private void OnTabRenaming(object sender, CancelEventArgs e)
	    {
			var e2 = new SheetTabRenamingEventArgs(canvas.Children.IndexOf(sender as UIElement));
			TabRenaming?.Invoke(this, e2);
			if(!e2.Cancel)
				IsInEditMode = true;
	                }


		private void OnTabRename(object sender, NamedEventArgs e)
	            {
			TabRename?.Invoke(this, new SheetTabRenameEventArgs(canvas.Children.IndexOf(sender as UIElement), e.Name));
	                }

		private void OnTabRenamed(object sender, EventArgs eventArgs)
	            {
			IsInEditMode = false;
			TabRenamed?.Invoke(this, new SheetTabRenamedEventArgs(canvas.Children.IndexOf(sender as UIElement)));
	            }
        
		#endregion
        
		public double TranslateScrollPoint(int p)
	    {
			return this.canvas.RenderTransform.Transform(new Point(p, 0)).X;
	    }
        
		public Rect GetItemBounds(int index)
	    {
			if (index < 0 || index > this.canvas.Children.Count - 1)
	    {
				throw new ArgumentOutOfRangeException("index");
	    }

			var tab = this.canvas.Children[index];

			return new Rect(tab.PointToScreen(new Point(0, 0)), this.RenderSize);
		}

		public void MoveItem(int index, int targetIndex)
	    {
			if (index < 0 || index > this.canvas.Children.Count - 1)
			{
				throw new ArgumentOutOfRangeException("index");
	    }

			var tab = this.canvas.Children[index];

			this.canvas.Children.RemoveAt(index);

			if (targetIndex > index) targetIndex--;

			this.canvas.Children.Insert(targetIndex, tab);
		}

		public void ScrollToItem(int index)
		{
			// TODO!

			//double width = this.ColumnDefinitions[2].ActualWidth;
			//int visibleWidth = this.ClientRectangle.Width - leftPadding - rightPadding;

			//if (rect.Width > visibleWidth || rect.Left < this.viewScroll + leftPadding)
			//{
			//	this.viewScroll = rect.Left - leftPadding;
			//}
			//else if (rect.Right - this.viewScroll > this.ClientRectangle.Right - rightPadding)
			//{
			//	this.viewScroll = rect.Right - this.ClientRectangle.Width + leftPadding;
			//}
		}

		public double ControlWidth { get; set; }

        public event EventHandler<SheetTabMovedEventArgs> TabMoved;

	    public event EventHandler SelectedIndexChanged;

	    public event EventHandler SplitterMoving;

	    public event EventHandler SheetListClick;

	    public event EventHandler NewSheetClick;

		public event EventHandler<SheetTabMouseEventArgs> TabMouseDown;
		
	    public event EventHandler<SheetTabRemovingEventArgs> SheetTabRemoving;

	    public event EventHandler<SheetTabRemoveEventArgs> SheetTabRemove;

	    public event EventHandler<SheetTabRenamingEventArgs> TabRenaming;

	    public event EventHandler<SheetTabRenameEventArgs> TabRename;

	    public event EventHandler<SheetTabRenamedEventArgs> TabRenamed;
    }

	sealed class SheetTabItem : Decorator
	{
        #region Приватные поля

        private Grid _grid;

	    private SheetTabItemHeader _label;

	    private Action<int> _renameSheetTabItemCallback;
		
        #endregion

		private SheetTabControl owner;

        public static readonly DependencyProperty IsSelectedProperty =
			DependencyProperty.Register("IsSelected", typeof(Boolean), typeof(SheetTabItem));

        public bool IsSelected
	    {
			get { return (bool)GetValue(IsSelectedProperty); }
	        set
	        {
	            bool currentValue = (bool)GetValue(IsSelectedProperty);

	            if (currentValue != value)
	            {
	                SetValue(IsSelectedProperty, value);
					this.InvalidateVisual();
	            }
	        }
	    }

		public static readonly DependencyProperty IsCanRemoveProperty =
			DependencyProperty.Register(nameof(IsCanRemove), typeof(Boolean), typeof(SheetTabItem));

	    public bool IsCanRemove
	    {
	        get => (bool)GetValue(IsCanRemoveProperty);
	        set
	        {
	            SetValue(IsCanRemoveProperty, value);
	            InvalidateVisual();
	        }
	    }

		public SheetTabItem(SheetTabControl owner, string title)
		{
			this.owner = owner;

			this.SnapsToDevicePixels = true;
			InitializeImpl();
			
			this.ChangeTitle(title);
		}

		public void ChangeTitle(string title)
		{
			// var label = new TextBlock
        // {
			// 	Text = title,
			// 	VerticalAlignment = System.Windows.VerticalAlignment.Center,
			// 	HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
			// 	Background = Brushes.Transparent,
			// };
			// 
			// label.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
			// 
			// this.Child = label;
			// this.Width = label.DesiredSize.Width + 9;
			_label.ChangeTitle(title);

			// _label.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
			_grid.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
			Width = _grid.DesiredSize.Width + 9;
		}

		private GuidelineSet gls = new GuidelineSet();

        public Color BackColor { get; set; }
	    public Color TextColor { get; set; }

        protected override void OnRender(DrawingContext drawingContext)
	    {
	        var g = drawingContext;

	        double right = RenderSize.Width;
	        double bottom = RenderSize.Height;

	        gls.GuidelinesX.Clear();
	        gls.GuidelinesY.Clear();

	        gls.GuidelinesX.Add(0.5);
	        gls.GuidelinesX.Add(right + 0.5);
	        gls.GuidelinesY.Add(0.5);
	        gls.GuidelinesY.Add(bottom + 0.5);

	        g.PushGuidelineSet(gls);

	        Brush b = new SolidColorBrush(owner.BorderColor);
	        var p = new Pen(b, 1);

	        if (IsSelected)
	        {
	            g.DrawRectangle(
					this.BackColor.A > 0 ? new SolidColorBrush(this.BackColor) : Brushes.White,
	                null, new Rect(0, 0, right, bottom));

	            g.DrawLine(p, new Point(0, 0), new Point(0, bottom));
	            g.DrawLine(p, new Point(right, 0), new Point(right, bottom));

	            g.DrawLine(p, new Point(0, bottom), new Point(right, bottom));

	            g.DrawLine(new Pen(Brushes.White, 1), new Point(1, 0), new Point(right, 0));
	        }
	        else
	        {
	            g.DrawRectangle(
					this.BackColor.A > 0 ? new SolidColorBrush(this.BackColor) : SystemColors.ControlBrush,
	                null, new Rect(0, 0, right, bottom));

				int index = this.owner.canvas.Children.IndexOf(this);

	            if (index > 0)
	            {
	                    g.DrawLine(new Pen(SystemColors.ControlDarkDarkBrush, 1), new Point(0, 2), new Point(0, bottom - 2));
	                }

	            // top border
	            g.DrawLine(p, new Point(0, 0), new Point(right, 0));
	        }

	        g.Pop();
	    }

		#region События

		public event EventHandler<CancelEventArgs> TabRenaming;

		public event EventHandler<NamedEventArgs> TabRename;

		public event EventHandler<EventArgs> TabRenamed;
		
		#endregion
		
        private void Image_MouseUp(object sender, MouseButtonEventArgs e)
        {
            owner.SheetTabRemoveRequest(this);
        }

        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs mouseButtonEventArgs)
	    {
		}
	        
		private void OnSizeChanging(object sender, EventArgs eventArgs)
		{
	    }

		private void OnTabRenaming(object o, CancelEventArgs e)
		{
			TabRenaming?.Invoke(this, e);
		}

		private void OnTabRename(object o, NamedEventArgs e)
		{
			TabRename?.Invoke(this, e);
		}

		private void OnTabRenamed(object sender, EventArgs e)
		{
			TabRenamed?.Invoke(this, e);
		}
		
        private void InitializeImpl()
	    {
	        _grid = new Grid();
	        _grid.ColumnDefinitions.Add(new ColumnDefinition { });
	        _grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });

            #region Title

	        var content = _label = new SheetTabItemHeader();
	        _label.TabRenaming += OnTabRenaming;
	        _label.TabRename += OnTabRename;
	        _label.TabRenamed +=OnTabRenamed;
            _label.SizeChanging += OnSizeChanging;

            Grid.SetColumn(content, 0);
	        

            #endregion

	        #region Image
        
            var imageSource = new BitmapImage();
	        imageSource.BeginInit();
	        imageSource.StreamSource = new System.IO.MemoryStream(Properties.Resources.close_png);
	        imageSource.EndInit();

	        var image = new Image
	        {
	            Source = imageSource,
	            VerticalAlignment = VerticalAlignment.Center,
	            HorizontalAlignment = HorizontalAlignment.Center,
	        };
	        image.SetBinding(IsEnabledProperty, new Binding(nameof(IsCanRemove))
	        {
	            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(SheetTabItem), 1),
	        });
	        Grid.SetColumn(image, 1);
	        #endregion

            _grid.Children.Add(content);
	        _grid.Children.Add(image);
	        image.MouseUp += Image_MouseUp;

	        Child = _grid;

	    }
	    }

	class ArrowBorder : Border
	{
		private SheetTabControl owner;
		
		public ArrowBorder(SheetTabControl owner)
		{
			this.owner = owner;
			
			SnapsToDevicePixels = true;
		}

		private GuidelineSet gls = new GuidelineSet() { GuidelinesY = new DoubleCollection(new double[] { 0.5 }) };

		protected override void OnRender(DrawingContext dc)
		{
			base.OnRender(dc);

			var g = dc;

			g.PushGuidelineSet(this.gls);

			g.DrawLine(new Pen(new SolidColorBrush(this.owner.BorderColor), 1),
				new Point(0, 0), new Point(this.RenderSize.Width, 0));

			g.Pop();
		}
	}

	class RightThumb : FrameworkElement
	{
		private SheetTabControl owner;

		public RightThumb(SheetTabControl owner)
		{
			this.owner = owner;
		}

		protected override Size MeasureOverride(Size availableSize)
		{
			return new Size(5, 0);
		}

		protected override void OnRender(DrawingContext drawingContext)
		{
			var g = drawingContext;

			var b = new SolidColorBrush(owner.BorderColor);
			var p = new Pen(b, 1);

			for (double y = 3; y < this.RenderSize.Height - 3; y += 4)
			{
				g.DrawRectangle(SystemColors.ControlDarkBrush, null, new Rect(0, y, 2, 2));
			}

			double right = this.RenderSize.Width;

			GuidelineSet gls = new GuidelineSet();
			gls.GuidelinesX.Add(right + 0.5);
			g.PushGuidelineSet(gls);

			g.DrawLine(p, new Point(right, 0), new Point(right, this.RenderSize.Height));

			g.Pop();
		}
	}

		
    class SheetTabItemHeader : ContentControl
    {
        #region Свойства зависимости

        public static readonly DependencyProperty IsInEditModeProperty =
            DependencyProperty.Register(nameof(IsInEditMode), typeof(Boolean), typeof(SheetTabItem));

        public static readonly DependencyProperty EditingTextProperty =
            DependencyProperty.Register(nameof(EditingText), typeof(string), typeof(SheetTabItem));
        #endregion

        #region Свойства

        public bool IsInEditMode
        {
            get => (bool)GetValue(IsInEditModeProperty);
            set => SetValue(IsInEditModeProperty, value);
        }

        public string EditingText
        {
            get => (string)GetValue(EditingTextProperty);
            set => SetValue(EditingTextProperty, value);
        }

        #endregion

        #region Приватные поля

        private TextBlock _label;

        #endregion

        #region Конструктор

        public SheetTabItemHeader()
        {
            Content = _label = new TextBlock
            {
                Text = string.Empty,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Background = Brushes.Transparent,
            };

            
            var trigger = new Trigger {Property = IsInEditModeProperty, Value = true};
            var dataTemplate = new DataTemplate();
            var factoryStackPanel = new FrameworkElementFactory(typeof(StackPanel));

            var factory = new FrameworkElementFactory(typeof(TextBox));
			factory.AddHandler(LostFocusEvent, new RoutedEventHandler(LostFocusHandler));
			factory.AddHandler(KeyUpEvent, new KeyEventHandler(KeyUpHandler));
			factory.AddHandler(GotFocusEvent, new RoutedEventHandler(GotFocusHandler));
			factory.AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler(TextChangedEventHandler));
            factory.SetValue(FocusExtension.IsFocusedProperty, true);
            factory.SetBinding(TextBox.TextProperty, new Binding(nameof(EditingText))
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(SheetTabItemHeader), 1),
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger =  UpdateSourceTrigger.PropertyChanged,
            });
			factory.SetValue(VerticalContentAlignmentProperty, VerticalAlignment.Stretch);

            //factoryStackPanel.AppendChild(factory);
            //dataTemplate.VisualTree = /*factory*/factoryStackPanel;
            dataTemplate.VisualTree = factory;
            var setter = new Setter(ContentTemplateProperty, dataTemplate);
            trigger.Setters.Add(setter);
            var style = new Style(typeof(SheetTabItemHeader));
            style.Triggers.Add(trigger);
            Style = style;
            
            MouseDoubleClick += MouseDoubleClickHandler;
        }


        #endregion

        #region Методы

        public void ChangeTitle(string title)
        {
            _label.Text = title;
            EditingText = title;
        }

        private void ChangeSizeRequest()
        {
            SizeChanging?.Invoke(this, new EventArgs());
        }

        #endregion

        #region События

        public event EventHandler<CancelEventArgs> TabRenaming;

        public event EventHandler<NamedEventArgs> TabRename;

        public event EventHandler<EventArgs> TabRenamed;

        public event EventHandler<EventArgs> SizeChanging;

        #endregion

        #region Обработчики событий

        private void MouseDoubleClickHandler(object sender, MouseButtonEventArgs e)
        {
            if (!IsInEditMode)
            {
                var renaming = new CancelEventArgs();
                TabRenaming?.Invoke(this, renaming);
                if (!renaming.Cancel)
                {
                    IsInEditMode = true;
                }
                else
                {
                    TabRenamed?.Invoke(this, new EventArgs());
                }
            }
        }

        private void LostFocusHandler(object sender, RoutedEventArgs e)
        {
            try
            {
                TabRename(this, new NamedEventArgs(EditingText));
            }
            finally
            {
                IsInEditMode = false;
                TabRenamed(this, new EventArgs());
            }
        }

        private void KeyUpHandler(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                IsInEditMode = false;
                TabRenamed(this, new EventArgs());
				e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                try
                {
                    TabRename(this, new NamedEventArgs(EditingText));
                }
                finally
                {
                    IsInEditMode = false;
                    TabRenamed(this, new EventArgs());
					e.Handled = true;
                }
            }
        }

        private void GotFocusHandler(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox txt)
            {
                txt.Text = EditingText;
                txt.SelectAll();
            }
        }

        private static void OnEditingTextPropertyChanged(
            DependencyObject d,
            DependencyPropertyChangedEventArgs e)
        {
            if (d is SheetTabItemHeader item)
            {
                item.ChangeSizeRequest();
            }
        }

        private void TextChangedEventHandler(object sender, TextChangedEventArgs e)
        {
            if (IsInEditMode)
            {
                ChangeSizeRequest();
            }
        }

        #endregion


    }

    class NamedEventArgs : EventArgs
    {
        public NamedEventArgs(string name)
        {
            Name = name;
        }

        public string Name { get; private set; }
    }

    public static class FocusExtension
    {
        public static bool GetIsFocused(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsFocusedProperty);
        }

        public static void SetIsFocused(DependencyObject obj, bool value)
        {
            obj.SetValue(IsFocusedProperty, value);
        }

        public static readonly DependencyProperty IsFocusedProperty =
            DependencyProperty.RegisterAttached(
                "IsFocused", typeof(bool), typeof(FocusExtension),
                new UIPropertyMetadata(false, OnIsFocusedPropertyChanged));


        private static void OnIsFocusedPropertyChanged(
            DependencyObject d,
            DependencyPropertyChangedEventArgs e)
        {
            var uie = (UIElement)d;
            if ((bool)e.NewValue)
            {
                uie.Focus(); // Don't care about false values.
            }
        }

    }
}

#endif // WPF