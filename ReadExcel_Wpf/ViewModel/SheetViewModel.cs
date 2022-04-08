using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WPF_Exercise.Model;

namespace WPF_ReadExcel.ViewModel
{
    public class SheetViewModel : Base
    {
        SheetModel Model { get; set; }
        public string Namesheet
        {
            get
            {
                return Model.Namesheet;
            }
            set
            {
                Model.Namesheet = value;
                OnPropertyChanged(nameof(Namesheet));
            }
        }
        public DataTable DataSheet
        {
            get { return Model.DataSheet; }
            set
            {
                Model.DataSheet = value;
                OnPropertyChanged(nameof(DataSheet));
            }
        }
        public SheetViewModel(SheetModel sheetModel)
        {           
            Model = sheetModel;
        }

    }
}
