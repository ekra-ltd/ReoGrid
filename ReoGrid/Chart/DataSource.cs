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
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using unvell.ReoGrid.Data;
using unvell.ReoGrid.Events;

namespace unvell.ReoGrid.Chart
{
    /// <summary>
    /// Represents the interface of data source used for chart.
    /// </summary>
    /// <typeparam name="T">Standard data serial classes.</typeparam>
    public interface IChartDataSource<T> : IDataSource<T> where T : IChartDataSerial
	{
		/// <summary>
		/// Get number of categories.
		/// </summary>
		int CategoryCount { get; }

		/// <summary>
		/// Get category name by specified index position.
		/// </summary>
		/// <param name="index">Zero-based number of category to get its name.</param>
		/// <returns>Specified category's name by index position.</returns>
		string GetCategoryName(int index);
	}

	/// <summary>
	/// Represents the interface of data serial used for chart.
	/// </summary>
	public interface IChartDataSerial : IDataSerial
	{
		/// <summary>
		/// Get the serial name.
		/// </summary>
		string Label { get; }
	}

/// <summary>
    /// Data source from given worksheet
    /// </summary>
    public class WorksheetChartDataSource : IChartDataSource<WorksheetChartDataSerial>
    {
        #region Constructor

        /// <summary>
        /// Create data source instance with specified worksheet instance
        /// </summary>
        /// <param name="worksheet">Instance of worksheet to read titles and data of plot serial.</param>
        public WorksheetChartDataSource(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        /// <summary>
        /// Create data source instance with specified worksheet instance
        /// </summary>
        /// <param name="worksheet">Instance of worksheet to read titles and data of plot serial.</param>
        /// <param name="serialNamesRange">Names for serial data from this range.</param>
        /// <param name="serialsRange">Serial data from this range.</param>
        /// <param name="serialPerRowOrColumn">Add serials by this specified direction. Default is Row.</param>
        public WorksheetChartDataSource(
            Worksheet worksheet,
            string serialNamesRange,
            string serialsRange,
            RowOrColumn serialPerRowOrColumn = RowOrColumn.Row)
            : this(worksheet)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (!worksheet.TryGetRangeByAddressOrName(serialNamesRange, out var snRange))
            {
                throw new InvalidAddressException("cannot determine the serial names range by specified range address or name.");
            }

            if (!worksheet.TryGetRangeByAddressOrName(serialsRange, out var sRange))
            {
                throw new InvalidAddressException("cannot determine the serials range by specified range address or name.");
            }

            AddSerialsFromRange(snRange, sRange, serialPerRowOrColumn);
        }

        /// <summary>
        /// Create data source instance with specified worksheet instance and serial data range.
        /// </summary>
        /// <param name="worksheet">Instance of worksheet to read titles and data of plot serial.</param>
        /// <param name="serialNamesRange">Range to read labels of data serial.</param>
        /// <param name="serialsRange">Range to read serial data.</param>
        /// <param name="serialPerRowOrColumn">Add serials by this specified direction. Default is Row.</param>
        public WorksheetChartDataSource(Worksheet worksheet, 
            RangePosition serialNamesRange, 
            RangePosition serialsRange,
            RowOrColumn serialPerRowOrColumn = RowOrColumn.Row)
            : this(worksheet)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            AddSerialsFromRange(serialNamesRange, serialsRange, serialPerRowOrColumn);
        }

        private void AddSerialsFromRange(
            RangePosition serialNamesRange,
            RangePosition serialsRange,
            RowOrColumn serialPerRowOrColumn = RowOrColumn.Row)
        {
            if (serialPerRowOrColumn == RowOrColumn.Row)
            {
                var offset = 0;
                for (var r = serialsRange.Row; r <= serialsRange.EndRow; r++, offset++)
                {
                    var labelPosiotion = new CellPosition(serialNamesRange.Row + offset, serialNamesRange.Col);
                    var label = new WorksheetedCellPosition(_worksheet, labelPosiotion);
                    var valuesPosition = new RangePosition(r, serialsRange.Col, 1, serialsRange.Cols);
                    AddSerial(label, new WorksheetedRangePosition(_worksheet, valuesPosition));
                }
            }
            else
            {
                var offset = 0;
                for (var c = serialsRange.Col; c <= serialsRange.EndCol; c++, offset++)
                {
                    var labelPosition = new CellPosition(serialNamesRange.Row, serialNamesRange.Col + offset);
                    var label = new WorksheetedCellPosition(_worksheet, labelPosition);
                    var valuesPosition = new RangePosition(serialsRange.Row, c, serialsRange.Rows, 1);
                    AddSerial(label, new WorksheetedRangePosition(_worksheet, valuesPosition));
                }
            }
        }

        #endregion // Constructor

        [Obsolete("Данное свойство ни имеет смысла")]
        /// <summary>
        /// Get instance of worksheet
        /// </summary>
        public Worksheet Worksheet { get; protected set; }

        #region Ranges

        /// <summary>
        /// Get or set the range to read row serial titles.
        /// </summary>
        [Obsolete("use AddSerial method to add data serial instead")]
        public virtual RangePosition RowTitleRange { get; set; }

        /// <summary>
        /// Get or set the range to read column serial titles.
        /// </summary>
        [Obsolete("use CategoryNameRange instead")]
        public virtual RangePosition ColTitleRange { get; set; }


        #endregion // Ranges

        #region Changes
        /// <summary>
        /// This method will be invoked when any data from the serial data range changed.
        /// </summary>
        public virtual void OnDataChanged()
        {
            DataChanged?.Invoke(this, null);
        }

        ///// <summary>
        ///// This method will be invoked when the serial data range changed.
        ///// </summary>
        //public virtual void OnDataRangeChanged()
        //{
        //	if (this.DataRangeChanged != null)
        //	{
        //		this.DataRangeChanged(this, null);
        //	}
        //}

        /// <summary>
        /// This event will be raised when data from the serial data range changed.
        /// </summary>
        public event EventHandler DataChanged;

        ///// <summary>
        ///// This event will be raised when the serial data range changed.
        ///// </summary>
        //public event EventHandler DataRangeChanged;

        #endregion // Changes

        #region Category

        /// <summary>
        /// Get or set the range that contains the category names.
        /// </summary>
        public WorksheetedRangePosition CategoryNameRange { get; set; }

        /// <summary>
        /// Return the title of specified column.
        /// </summary>
        /// <param name="index">Zero-based number of column.</param>
        /// <returns>Return the title that will be displayed on chart.</returns>
        public string GetCategoryName(int index)
        {
            if (CategoryNameRange?.Position.IsEmpty != false)
            {
                return null;
            }
            var worksheet = CategoryNameRange?.Worksheet ?? _worksheet;
            if(CategoryNameRange.Position.Cols == 1 && CategoryNameRange.Position.Rows > 1)
                return worksheet.GetCellData<string>(CategoryNameRange.Position.Row + index, CategoryNameRange.Position.Col );
            return worksheet.GetCellData<string>(CategoryNameRange.Position.Row, CategoryNameRange.Position.Col + index);
        }

        #endregion // Category

        #region Serials
        /// <summary>
        /// Get number of data serials.
        /// </summary>
        public virtual int SerialCount => serials.Count;

        /// <summary>
        /// Get number of categories.
        /// </summary>
        public virtual int CategoryCount => _categoryCount;

        internal List<WorksheetChartDataSerial> serials = new List<WorksheetChartDataSerial>();

        public WorksheetChartDataSerial this[int index]
        {
            get
            {
                if (index < 0 || index >= serials.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }

                return serials[index];
            }
        }

        /// <summary>
        /// Add serial data into data source.
        /// </summary>
        /// <param name="serial">Serial data source.</param>
        public void Add(WorksheetChartDataSerial serial)
        {
            serials.Add(serial);

            UpdateCategoryCount(serial);
        }

        internal void UpdateCategoryCount(WorksheetChartDataSerial serial)
        {
            if (serial.DataRange is null)
            {
                Debug.Fail($"Ожидается, что {nameof(serial)}.{nameof(serial.DataRange)} будет не null");
                return;
            }
            _categoryCount = Math.Max(_categoryCount, Math.Max(serial.DataRange.Position.Cols, serial.DataRange.Position.Rows));
        }

        /// <summary>
        /// Add serial data into data source from a range, set the name as the label of serial.
        /// </summary>
        /// <param name="name">Name for serial to be added.</param>
        /// <param name="serialRange">Range to read serial data from worksheet.</param>
        /// <returns>Instance of chart serial has been added.</returns>
        public WorksheetChartDataSerial AddSerial(WorksheetedCellPosition labelAddress, WorksheetedRangePosition serialRange)
        {
            var serial = new WorksheetChartDataSerial(this, labelAddress, serialRange);
            Add(serial);
            return serial;
        }

        public WorksheetChartDataSerial GetSerial(int index)
        {
            if (index < 0 || index >= serials.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return serials[index];
        }
        
        /// <summary>
        /// Get collection of data serials.
        /// </summary>
        public WorksheetChartDataSerialCollection Serials 
            => _collection ?? (_collection = new WorksheetChartDataSerialCollection(this));

        #endregion // Serials

        #region Приватные поля

        private readonly Worksheet _worksheet;

        private WorksheetChartDataSerialCollection _collection;

        private int _categoryCount;

        #endregion
    }

#region WorksheetChartDataSerialCollection
	/// <summary>
	/// Represents collection of data serial.
	/// </summary>
	public class WorksheetChartDataSerialCollection : IList<WorksheetChartDataSerial>
	{
		public WorksheetChartDataSource DataSource { get; private set; }

		private List<WorksheetChartDataSerial> serials;

		internal WorksheetChartDataSerialCollection(WorksheetChartDataSource dataSource)
		{
			this.DataSource = dataSource;
			this.serials = dataSource.serials;
		}

		public WorksheetChartDataSerial this[int index]
		{
			get { return this.serials[index]; }
			set
			{
				this.serials[index] = value;
				this.DataSource.UpdateCategoryCount(value);
			}
		}

		public int Count
		{
			get { return this.serials.Count;}
		}

		public bool IsReadOnly { get { return false; } }

		public void Add(WorksheetChartDataSerial serial)
		{
			this.DataSource.Add(serial);
		}

		public void Clear()
		{
			this.serials.Clear();
			// TODO: update category count
		}

		public bool Contains(WorksheetChartDataSerial item)
		{
			return this.serials.Contains(item);
		}

		public void CopyTo(WorksheetChartDataSerial[] array, int arrayIndex)
		{
			this.serials.CopyTo(array, arrayIndex);
		}

		public IEnumerator<WorksheetChartDataSerial> GetEnumerator()
		{
			return this.serials.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return this.serials.GetEnumerator();
		}

		public int IndexOf(WorksheetChartDataSerial item)
		{
			return this.serials.IndexOf(item);
		}

		public void Insert(int index, WorksheetChartDataSerial serial)
		{
			this.serials.Insert(index, serial);

			this.DataSource.UpdateCategoryCount(serial);
		}

		public bool Remove(WorksheetChartDataSerial serial)
		{
			return this.serials.Remove(serial);
			// TODO: update category count
		}

		public void RemoveAt(int index)
		{
			this.serials.RemoveAt(index);
			// TODO: update category count
		}
	}
	#endregion // WorksheetChartDataSerialCollection

	#region WorksheetChartDataSerial
	
	/// <summary>
    /// Represents implementation of chart data serial.
    /// </summary>
    // ReSharper disable once InheritdocConsiderUsage
    public class WorksheetChartDataSerial : IChartDataSerial
    {
        #region Конструктор

        /// <param name="dataSource">Data source to read chart data from worksheet.</param>
        /// <param name="worksheet">Instance of worksheet that contains the data to be read.</param>
        /// <param name="labelAddress">The address to locate label of serial on worksheet.</param>
        /// <param name="dataRange">Serial data range to read serial data for chart from worksheet.</param>
        public WorksheetChartDataSerial(
            WorksheetChartDataSource dataSource,
            // Worksheet worksheet,
            WorksheetedCellPosition labelAddress,
            WorksheetedRangePosition dataRange)
            : this(dataSource, /*worksheet,*/ labelAddress)
        {
            _dataRange = dataRange;
        }

        /// <summary>
        /// Create data serial by specified worksheet instance and data range.
        /// </summary>
        /// <param name="dataSource">Data source to read chart data from worksheet.</param>
        /// <param name="worksheet">Instance of worksheet that contains the data to be read.</param>
        /// <param name="labelAddress">The address to locate label of serial on worksheet.</param>
        /// <param name="addressOrName">Serial data specified by address position or range's name.</param>
        public WorksheetChartDataSerial(
            WorksheetChartDataSource dataSource,
            Worksheet worksheet,
            string labelAddress,
            string addressOrName)
            : this(dataSource, /*worksheet,*/ new WorksheetedCellPosition(worksheet, labelAddress))
        {
            if (RangePosition.IsValidAddress(addressOrName))
            {
                _dataRange = new WorksheetedRangePosition(worksheet, addressOrName);
            }
            else if (NamedRange.IsValidName(addressOrName))
            {
                if (worksheet != null)
                {
                    if (worksheet.TryGetNamedRange(addressOrName, out var range))
                    {
                        _dataRange = new WorksheetedRangePosition(worksheet, range);
                    }
                    else
                    {
                        throw new InvalidAddressException(addressOrName);
                    }
                }
                else
                {
                    throw new ReferenceObjectNotAssociatedException("Data source must associate to valid worksheet instance.");
                }
            }
            else
            {
                throw new InvalidAddressException(addressOrName);
            }
        }

        protected WorksheetChartDataSerial(WorksheetChartDataSource dataSource, /*Worksheet worksheet,*/ WorksheetedCellPosition labelAddress)
        {
            if (dataSource == null)
            {
                throw new ArgumentNullException(nameof(dataSource));
            }
            _dataSource = dataSource;
            LabelAddress = labelAddress;
            AddEventHandlers();
        }



        #endregion

        #region Деструктор

        /// <summary>
        /// Destroy the worksheet data serial and release all event handlers to data source.
        /// </summary>
        ~WorksheetChartDataSerial()
        {
            RemoveEventHandlers();
        }

        #endregion

        #region Методы

        private void AddEventHandlers()
        {
            if (LabelAddress?.Worksheet != null)
            {
                AddCellDataChangedEventHandlers(LabelAddress.Worksheet);
                RangeDataChangedEventHandlers(LabelAddress.Worksheet);
            }
            if (DataRange?.Worksheet != null)
            {
                AddCellDataChangedEventHandlers(DataRange.Worksheet);
                RangeDataChangedEventHandlers(DataRange.Worksheet);
            }
        }

        private void AddCellDataChangedEventHandlers(Worksheet worksheet)
        {
            if (worksheet is null) return;
            if (!_cellDataChangedHanlders.ContainsKey(worksheet))
            {
                _cellDataChangedHanlders[worksheet] = (sender, args) =>
                {
                    worksheet_CellDataChanged(args, worksheet);
                };
                worksheet.CellDataChanged +=  _cellDataChangedHanlders[worksheet];
            }
        }

        private void RangeDataChangedEventHandlers(Worksheet worksheet)
        {
            if (worksheet is null) return;
            if (!_rangeEventArgsDataChangedHanlders.ContainsKey(worksheet))
            {
                _rangeEventArgsDataChangedHanlders[worksheet] = (sender, args) =>
                {
                    worksheet_RangeDataChanged(args, worksheet);
                };
                worksheet.RangeDataChanged += _rangeEventArgsDataChangedHanlders[worksheet];
            }
        }

        private void RemoveEventHandlers()
        {
            foreach (var worksheet in _cellDataChangedHanlders.Keys)
            {
                try
                {
                    worksheet.CellDataChanged -= _cellDataChangedHanlders[worksheet];
                }
                catch (Exception exc)
                {
                    Debug.Fail(exc.Message);
                    // ignore
                }
            }
            foreach (var worksheet in _rangeEventArgsDataChangedHanlders.Keys)
            {
                try
                {
                    worksheet.RangeDataChanged -= _rangeEventArgsDataChangedHanlders[worksheet];
                }
                catch (Exception exc)
                {
                    Debug.Fail(exc.Message);
                    // ignore
                }
            }
        }

        #endregion

        #region Свойства


        /// <summary>
        /// Get instance of worksheet
        /// </summary>
        [Obsolete("Данное свойство ни имеет смысла")]
        public Worksheet Worksheet { get; protected set; }


        /// <summary>
        /// Determine the range to read data from worksheet
        /// </summary>
        public virtual WorksheetedRangePosition DataRange
        {
            get => _dataRange;
            set
            {
                if (_dataRange != value)
                {
                    _dataRange = value;

                    //this.dataSource.OnDataRangeChanged();
                    _dataSource.OnDataChanged();
                }
            }
        }

        //private string name;
        public WorksheetedCellPosition LabelAddress { get; set; }

        /// <summary>
        /// Get label text of serial.
        /// </summary>
        public string Label => LabelAddress.GetCellText();

        /// <summary>
        /// Get number of data items of current serial.
        /// </summary>
        // ReSharper disable once InheritdocConsiderUsage
        public int Count
        {
            get
            {
                if (_dataRange is null)
                {
                    Debug.Fail($"Ожидается что {nameof(_dataRange)} будет заполнен");
                    return 0;
                }
                var result = 0;
                if (DataRange.Position.Rows > DataRange.Position.Cols)
                {
                    result = DataRange.Position.Rows;
                }
                else
                {
                    result = DataRange.Position.Cols;
                }
                return result;
            }
        }

        /// <summary>
        /// Get data from serial by specified index position.
        /// </summary>
        /// <param name="index">Zero-based index position in serial to get data.</param>
        /// <returns>Data in double type to be get from specified index of serial.
        /// If index is out of range, or data in worksheet is null, then return null.
        /// </returns>
        // ReSharper disable once InheritdocConsiderUsage
        public double? this[int index]
        {
            get
            {
                if (DataRange is null)
                {
                    Debug.Fail($"Ожидается что {nameof(DataRange)} будет заполнен");
                    return null;
                }
                if (DataRange.Worksheet is null)
                {
                    Debug.Fail($"Ожидается что {nameof(DataRange)}.{nameof(DataRange.Worksheet)} будет заполнен");
                    return null;
                }


                object data;
                if (_dataRange.Position.Rows > _dataRange.Position.Cols)
                {
                    data = DataRange.Worksheet.GetCellData(_dataRange.Position.Row + index, _dataRange.Position.Col);
                }
                else
                {
                    data = DataRange.Worksheet.GetCellData(_dataRange.Position.Row, _dataRange.Position.Col + index);
                }

                if (Utility.CellUtility.TryGetNumberData(data, out var val))
                {
                    return val;
                }
                return null;
            }
        }

        #endregion

        #region Events

        void worksheet_CellDataChanged(CellEventArgs e, Worksheet worksheet)
        {
            var cell = e.Cell;
            if (DataRange != null)
            {
                if (worksheet == DataRange.Worksheet)
                {
                    if(DataRange.Position.Contains(cell.Position))
                        _dataSource.OnDataChanged();
                }
                if (_dataSource.CategoryNameRange != null)
                {
                    if (worksheet == _dataSource.CategoryNameRange.Worksheet)
                    {
                        if (_dataSource.CategoryNameRange.Position.Contains(cell.Position))
                            _dataSource.OnDataChanged();
                    }
                }
            }
            else if (LabelAddress != null)
            {
                if (worksheet == LabelAddress.Worksheet)
                {
                    if (LabelAddress.Position == cell.Position)
                        _dataSource.OnDataChanged();
                }
            }
            else
            {
                Debug.Fail($"Ожидается что {nameof(_dataRange)} будет заполнен");
            }
        }

        void worksheet_RangeDataChanged(RangeEventArgs e, Worksheet worksheet)
        {
            var range = e.Range;
            if (DataRange != null)
            {
                if (worksheet == DataRange.Worksheet)
                {
                    if (DataRange.Position.IntersectWith(range))
                        _dataSource.OnDataChanged();
                }
            }
            else if (LabelAddress != null)
            {
                if (worksheet == LabelAddress.Worksheet)
                {
                    if (range.Contains(LabelAddress.Position))
                        _dataSource.OnDataChanged();
                }
            }
            else
            {
                Debug.Fail($"Ожидается что {nameof(_dataRange)} будет заполнен");
            }
        }


        #endregion // Events

        #region Приватные поля

        private readonly WorksheetChartDataSource _dataSource;

        private WorksheetedRangePosition _dataRange;

        private readonly Dictionary<Worksheet, EventHandler<CellEventArgs>> _cellDataChangedHanlders = new Dictionary<Worksheet, EventHandler<CellEventArgs>>();

        private readonly Dictionary<Worksheet, EventHandler<RangeEventArgs>> _rangeEventArgsDataChangedHanlders = new Dictionary<Worksheet, EventHandler<RangeEventArgs>>();

        #endregion
    }
	
	#endregion // WorksheetChartDataSerial

}

#endif // DRAWING