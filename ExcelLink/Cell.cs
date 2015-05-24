using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;

namespace ExcelLink
{
    public class Cell : DependencyObject, INotifyPropertyChanged
    {
        public static DependencyProperty ValueProperty = 
            DependencyProperty.Register("Value", typeof(String), typeof(Cell),
                                        new PropertyMetadata(new PropertyChangedCallback(OnValueChanged)));

        private Sheet _sheet;
        
        private int _row;
        public int Row { get { return _row; } }

        private int _column;
        public int Column { get { return _column; } }
        
        public String Value
        {
            get
            {
                return GetValue(ValueProperty) as String;
            }

            set
            {
                SetValue(ValueProperty, value);
                RaisePropertyChanged("Value");
            }
        }

        private bool valueChangeFromExcel = false;
        private  static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            Cell c = d as Cell;

            // if the update comes from excel, we do not need to update excel again :)
            if (!c.valueChangeFromExcel)
            {
                c._sheet.SetValue(c._row, c._column, e.NewValue as String);
            }
        }
        
        public Cell(Sheet sh, int row, int column)
        {
            _sheet = sh;
            _row = row;
            _column = column;

            RaisePropertyChanged("Row");
            RaisePropertyChanged("Column");
            Reload();
        }

        public delegate void MethodInvoker();
        public void Reload()
        {
            this.Dispatcher.BeginInvoke((MethodInvoker) delegate 
            {
                // Reload is an update that comes from Excel (reading the value from the Excel sheet)
                // make sure the update does not get set back on the sheet (creating a loop)
                valueChangeFromExcel = true;
                SetValue(ValueProperty, _sheet.GetValue(_row, _column));
                valueChangeFromExcel = false;
                RaisePropertyChanged("Value");
            });
        }

        void RaisePropertyChanged(string prop) 
        { 
            if (PropertyChanged != null) 
            { 
                PropertyChanged(this, new PropertyChangedEventArgs(prop)); 
            } 
        } 
        public event PropertyChangedEventHandler PropertyChanged; 
    }
}
