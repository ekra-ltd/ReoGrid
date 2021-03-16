using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid.Chart
{
    /// <summary>
    /// Область легенды для круговой диаграммы
    /// </summary>
    /// <remarks>
    /// Поведение по умолчанию занимает всю область диаграммы, если рядов данных много, в результате самой диаграммы
    /// не видно - есть только легенда. Данный класс призван ограничить занимаемую облать накоторой частью диаграммы.
    /// Изначально 50% по высоте от области
    /// </remarks>
    public class PieChartLegend2 : ChartLegend
    {
        private const double HeightFraction = 0.5;

        public PieChartLegend2(IChart chart) : base(chart)
        {
        }

        #region Переопределенные методы

        public override void MeasureSize(Rectangle parentClientRect)
        {
            base.MeasureSize(FlattenRectangle(parentClientRect, HeightFraction));
        }

        #endregion

        private static Rectangle FlattenRectangle(Rectangle rect, double remainingFraction)
        {
            var flattenH = rect.Height * remainingFraction;
            return new Rectangle {X = rect.X, Y = rect.Y - flattenH, Height = flattenH, Width = rect.Width};
        }
    }
}