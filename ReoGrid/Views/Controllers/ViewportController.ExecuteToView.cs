using System;
using System.Linq;
using System.Collections.Generic;

namespace unvell.ReoGrid.Views
{
    internal partial class ViewportController
    {
        private bool ExecuteToView(IEnumerable<IView> views, Func<IView, bool> function)
        {
            return views.Any(subView =>
            {
                var result = false;
                if(subView != null)
                    result = function?.Invoke(subView) ?? false;
                return result;
            });
        }

    }
}