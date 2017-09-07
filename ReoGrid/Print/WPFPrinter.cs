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

#if PRINT

#if WPF

using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using unvell.ReoGrid.Print;
using unvell.ReoGrid.Rendering;

namespace unvell.ReoGrid.Print
{
	partial class PrintSession
	{
		internal void Init() { }

		public void Dispose() { }

		/// <summary>
		/// Start output document to printer.
		/// </summary>
		public void Print()
		{
			throw new NotImplementedException("WPF Print is not implemented yet. Try use Windows Form version to print document as XPS file.");
		}
	}

	internal class WorksheetPaginator: DocumentPaginator
	{
		private Worksheet _worksheet = null;

		private List<DocumentPage> _pages = new List<DocumentPage>();

		private Size _pageSize;

		private bool _isPageCountValid = false;
		public WorksheetPaginator(Worksheet worksheet)
		{
			_isPageCountValid = false;
			_worksheet = worksheet;
			PageSize = new Size(210 , 297);
		}
		public override DocumentPage GetPage(int pageNumber)
		{
			return _pages[pageNumber];
		}

		public override bool IsPageCountValid => _isPageCountValid;

		public override int PageCount => _pages.Count;

		public override Size PageSize
		{
			get { return _pageSize; }
			set
			{
				_isPageCountValid = false;
				_pageSize = value;
				PaginateData();
			}
		}

		public override IDocumentPaginatorSource Source => null;

		protected void PaginateData()
		{
			_worksheet.PrintSettings.PaperHeight = 11.03f;//11.69f;
			_worksheet.PrintSettings.PaperWidth = 7.8f;//8.27f;
			_worksheet.PrintSettings.Margins = new PageMargins(1/2.548, 1/2.548, 2/2.548, 1/2.548);
			_pages.Clear();
			var printEnable = _worksheet.BeginPrinting();
			if (printEnable)
			{
				try
				{
					while (printEnable)
					{
						DrawingVisual drawingVisual = new DrawingVisual();
						using (System.Windows.Media.DrawingContext drawingContext = drawingVisual.RenderOpen())
						{
							printEnable = _worksheet.ToImage(drawingContext);
						}
						var size = new Size(PageSize.Width * 780 / 210, PageSize.Height * 780 / 210);
						if (_worksheet.PrintSettings.Landscape)
						{
							size = new Size(PageSize.Height * 780 / 210, PageSize.Width * 780 / 210);
						}
						_pages.Add(new DocumentPage(drawingVisual, size, new Rect(size), new Rect(size)));
					}
				}
				finally
				{
					_worksheet.EndPrinting();
				}
			}
			_isPageCountValid = true;
		}
	}
}

namespace unvell.ReoGrid
{
	partial class Worksheet
	{
		internal bool ToImage(System.Windows.Media.DrawingContext graphics)
		{
			var renderer = new WPFRenderer {PlatformGraphics = graphics};

			CellDrawingContext context = new CellDrawingContext(this, DrawMode.Print, renderer)
			{
				Graphics = renderer,
				AllowCellClip = true,
			};

			var saveContext = _printSession.DrawingContext;
			try
			{
				_printSession.DrawingContext = context;
				_printSession.NextPage(graphics);
				return _printSession.hasMorePages;
			}
			finally
			{
				_printSession.DrawingContext = saveContext;
			}
		}

		internal bool BeginPrinting()
		{
			_printSession = CreatePrintSession();
			
			if (_printSession != null)
			{
				_printSession.CurrentWorksheetIndex = -1;
				_printSession.NextWorksheet();
				return true;
			}
			return false;
		}

		internal void EndPrinting()
		{
			_printSession.Dispose();
			_printSession = null;
		}

		private PrintSession _printSession = null;

		public DocumentPaginator GetDocumentPaginator()
		{
			return new WorksheetPaginator(this);
		}

	}
}

#endif // WPF

#endif // PRINT