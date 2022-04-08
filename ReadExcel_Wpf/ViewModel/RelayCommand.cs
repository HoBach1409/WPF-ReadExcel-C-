using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace WPF_ReadExcel.ViewModel
{
    public class RelayCommand : ICommand //Bind command của control -> logic
    {
        readonly Action<object> execute; //k trả về 
        readonly Predicate<object> Canexecute; //trả về bool

        public Action<DataGrid> CanExportExcel { get; }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public RelayCommand(Action<object> ex, Predicate<object> canexecute)
        {
            execute = ex;
            Canexecute = canexecute;
        }
        public RelayCommand(Action<object> execute) : this(execute, null)
        {

        }

        public RelayCommand(Action<DataGrid> canExportExcel)
        {
            CanExportExcel = canExportExcel;
        }

        public bool CanExecute(object parameter)
        {
            return Canexecute == null ? true : Canexecute.Invoke(parameter);
        }

        public void Execute(object parameter)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //if(CanExecute(parameter))
                if (execute != null)
                {
                    execute.Invoke(parameter);
                }
                else
                {
                    throw new InvalidOperationException("Execute methor not find");
                }
        }
    }
}
