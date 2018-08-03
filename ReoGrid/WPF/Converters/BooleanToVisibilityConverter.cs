using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace unvell.ReoGrid.WPF.Converters
{
    public class BooleanToVisibilityConverter: IValueConverter
    {
        public Visibility FalseVisibility = Visibility.Collapsed;

        public Visibility TrueVisibility = Visibility.Visible;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool f)
            {
                return f ? TrueVisibility : FalseVisibility;
            }
            return DependencyProperty.UnsetValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
