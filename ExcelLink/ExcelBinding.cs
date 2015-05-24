using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows;

namespace ExcelLink
{
    public static class ExcelBinding
    {
        public static BindingExpressionBase Bind(Cell c, Object source, String propertyPath, BindingMode bindingMode = BindingMode.TwoWay, bool loadFromTarget = true)
        {

            if (loadFromTarget)
            {
                // bind one way to source to read from the target to the property path
                Binding b = new Binding(propertyPath);
                b.Source = source;
                b.Mode = BindingMode.OneWayToSource;
                BindingOperations.SetBinding(c, Cell.ValueProperty, b);
                BindingOperations.ClearBinding(c, Cell.ValueProperty);
            }

            Binding bind = new Binding(propertyPath);
            bind.Source = source;
            bind.Mode = bindingMode;

            return BindingOperations.SetBinding(c, Cell.ValueProperty, bind);
        }
    }
}
