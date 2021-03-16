using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using unvell.ReoGrid.Chart;

namespace unvell.ReoGrid.WPFDemo.Task
{
    /// <summary>
    /// Класс для тестирования задачи #9506
    /// Создает графики, используемые в scada в различных вариантах и выполняет демнострацию их работы
    /// </summary>
    internal class Task9506HorizontalHistogram : ITaskExample
    {
        #region Публичные методы

        #region Interface ITaskExample

        public void Apply(ReoGridControl grid)
        {
            #region Лист с данными и графиками

            var worksheet = grid.NewWorksheet("#9506");
            var generator = new CombinedGenerator(
                new SinValueGenerator(100, 1 / 200.0f, 10),
                new SinValueGenerator(10, 1 / 20.0f, 10));

            var rangesByColumn = AddValues(worksheet, generator, ValuesCount, new CellPosition("A1"), ValuesFillDirection.Columns);
            var rangesByRows = AddValues(worksheet, generator, ValuesCount, new CellPosition("A5"), ValuesFillDirection.Rows);


            double scale = 2;
            var firstChartRect = new Rect(new Point(430, 210), new Size(scale * 200, scale * 150));


            var chartsAdders = new AddChartDelegate[]
            {
                AddCharts<LineChart>, // Линейный график
                AddCharts<ColumnChart>, // Гистограмма
                AddCharts<Pie2DChart>, // Круговая диаграмма
                AddCharts<BarChart>, // Горизонтальная гистограмма
                AddCharts<AreaChart> // Диаграмма с областями

                // Не используется в скада
                // AddCharts<DoughnutChart>(worksheet, firstChartRect, rangesByColumn, rangesByRows); 
                // AddCharts<PieChart>(worksheet, firstChartRect, rangesByColumn, rangesByRows);
            };

            foreach (var adder in chartsAdders)
            {
                adder(worksheet, firstChartRect, rangesByColumn, rangesByRows);
                MoveDown(ref firstChartRect);
            }

            #endregion
        }

        #endregion

        #endregion

        #region Вспомогательные методы

        private static void MoveDown(ref Rect rect)
        {
            rect.Y += rect.Height;
        }

        private static void MoveRight(ref Rect rect)
        {
            rect.X += rect.Width;
        }

        private static void AddCharts<T>(Worksheet worksheet, Rect firstChartRect, AddValuesResult rangesByColumn, AddValuesResult rangesByRows)
            where T : Chart.Chart, new()
        {
            AddChart<T>(worksheet, ref firstChartRect, rangesByColumn, RowOrColumn.Row);
            AddChart<T>(worksheet, ref firstChartRect, rangesByRows, RowOrColumn.Row);
            AddChart<T>(worksheet, ref firstChartRect, rangesByColumn, RowOrColumn.Column);
            AddChart<T>(worksheet, ref firstChartRect, rangesByRows, RowOrColumn.Column);
        }

        private static void AddChart<T>(Worksheet worksheet,
            ref Rect firstChartRect,
            AddValuesResult source,
            RowOrColumn dataSource)
            where T : Chart.Chart, new()
        {
            worksheet.FloatingObjects.Add(new T
            {
                Location = new Point(firstChartRect.X, firstChartRect.Y),
                Width = firstChartRect.Width,
                Height = firstChartRect.Height,

                Title = GetName(typeof(T).Name, ValuesCount, ValuesFillDirection.Rows, RowOrColumn.Column),
                DataSource = new WorksheetChartDataSource(worksheet, source.NameRange, source.DataRange, dataSource)
                {
                    CategoryNameRange = new WorksheetedRangePosition(worksheet, source.CatRange)
                },
                ShowLegend = true
            });
            MoveRight(ref firstChartRect);
        }

        private AddValuesResult AddValues(Worksheet worksheet, ValueGenerator generator, int count, CellPosition startPos, ValuesFillDirection direction)
        {
            // Данные в столбцах
            if (direction == ValuesFillDirection.Columns)
            {
                worksheet.ColumnCount = Math.Max(startPos.Col + ValuesCount, worksheet.ColumnCount);

                var nameFirstCellPos = new CellPosition(startPos.Row, startPos.Col);
                var catFirstCellPos = new CellPosition(startPos.Row + 1, startPos.Col);
                var dataFirstCellPos = new CellPosition(startPos.Row + 2, startPos.Col);

                // Заполнение данных
                for (var i = 0; i < ValuesCount; i++)
                {
                    worksheet[catFirstCellPos.Row, catFirstCellPos.Col + i] = i;

                    var rowCounter = 0;
                    foreach (var value in generator.Value(i))
                    {
                        worksheet[dataFirstCellPos.Row + rowCounter, dataFirstCellPos.Col + i] = value;
                        ++rowCounter;
                    }
                }

                // Построение chart-а
                worksheet[nameFirstCellPos] = ValuesFillDirectionText(direction);

                return new AddValuesResult
                {
                    DataRange = new RangePosition(dataFirstCellPos) {Cols = ValuesCount, Rows = generator.Value(0).Length},
                    CatRange = new RangePosition(catFirstCellPos) {Cols = ValuesCount},
                    NameRange = new RangePosition(nameFirstCellPos)
                };
            }

            if (direction == ValuesFillDirection.Rows)
            {
                worksheet.RowCount = Math.Max(startPos.Row + ValuesCount, worksheet.RowCount);

                var nameFirstCellPos = new CellPosition(startPos.Row, startPos.Col);
                var catFirstCellPos = new CellPosition(startPos.Row, startPos.Col + 1);
                var dataFirstCellPos = new CellPosition(startPos.Row, startPos.Col + 2);
                ;

                // Заполнение данных
                for (var i = 0; i < ValuesCount; i++)
                {
                    worksheet[catFirstCellPos.Row + i, catFirstCellPos.Col] = i;
                    //
                    var colCounter = 0;
                    foreach (var value in generator.Value(i))
                    {
                        worksheet[dataFirstCellPos.Row + i, dataFirstCellPos.Col + colCounter] = value;
                        ++colCounter;
                    }
                }

                // Построение chart-а
                worksheet[nameFirstCellPos] = ValuesFillDirectionText(direction);

                return new AddValuesResult
                {
                    DataRange = new RangePosition(dataFirstCellPos) {Rows = ValuesCount, Cols = generator.Value(0).Length},
                    CatRange = new RangePosition(catFirstCellPos) {Rows = ValuesCount},
                    NameRange = new RangePosition(nameFirstCellPos)
                };
            }

            throw new ArgumentException(nameof(direction));
        }

        private static string ValuesFillDirectionText(ValuesFillDirection d)
        {
            switch (d)
            {
                case ValuesFillDirection.Columns:
                    return @"Столбцы";
                case ValuesFillDirection.Rows:
                    return @"Строки";
                default:
                    throw new ArgumentOutOfRangeException(nameof(d), d, null);
            }
        }

        private static string GetName(string chartName, int count, ValuesFillDirection values, RowOrColumn dataDirection)
        {
            string ToString(RowOrColumn d)
            {
                switch (d)
                {
                    case RowOrColumn.Row:
                        return @"Строки";
                    case RowOrColumn.Column:
                        return @"Столбцы";
                    case RowOrColumn.Both:
                        return @"Строки+Столбцы";
                    default:
                        throw new ArgumentOutOfRangeException(nameof(d), d, null);
                }
            }

            return $"{chartName},{count},{ValuesFillDirectionText(values)},{ToString(dataDirection)}";
        }

        #endregion

        #region Вспомогательные типы

        private enum ValuesFillDirection
        {
            Columns,
            Rows
        }

        private struct AddValuesResult
        {
            public RangePosition DataRange { get; set; }
            public RangePosition CatRange { get; set; }
            public RangePosition NameRange { get; set; }
        }

        private abstract class ValueGenerator
        {
            #region Публичные методы

            #region Переопределяемые методы

            public abstract IEnumerable<double[]> Values(int count);

            public abstract double[] Value(int point);

            #endregion

            #endregion
        }

        private sealed class SinValueGenerator : ValueGenerator
        {
            #region Конструктор

            public SinValueGenerator(double amplitude, double phase, double offset)
            {
                _amplitude = amplitude;
                _phase = phase;
                _offset = offset;
            }

            #endregion

            #region Публичные методы

            public override IEnumerable<double[]> Values(int count)
            {
                for (var i = 0; i < count; ++i)
                    yield return Value(i);
            }

            public override double[] Value(int point)
            {
                return new[] {_amplitude * Math.Sin(2 * Math.PI * point * _phase) + _offset};
            }

            #endregion

            #region Поля

            private readonly double _amplitude;
            private readonly double _phase;
            private readonly double _offset;

            #endregion
        }

        private sealed class CombinedGenerator : ValueGenerator
        {
            #region Конструктор

            public CombinedGenerator(params ValueGenerator[] generators)
            {
                _generators = generators;
            }

            #endregion

            #region Публичные методы

            public override IEnumerable<double[]> Values(int count)
            {
                for (var i = 0; i < count; ++i)
                    yield return Value(i);
            }

            public override double[] Value(int point)
            {
                return _generators.SelectMany(g => g.Value(point)).ToArray();
            }

            #endregion

            #region Поля

            private readonly ValueGenerator[] _generators;

            #endregion
        }

        private delegate void AddChartDelegate(Worksheet worksheet, Rect firstChartRect, AddValuesResult rangesByColumn, AddValuesResult rangesByRows);

        #endregion

        private const int ValuesCount = 48; // ограничение по количеству колонок ~16к, ограничение на количество рядов данных = 255 (не откроется в excel)
    }
}