/*
The MIT License (MIT)

Copyright (c) 2015 Sander Cox - Parallel Dimension

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 */

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
