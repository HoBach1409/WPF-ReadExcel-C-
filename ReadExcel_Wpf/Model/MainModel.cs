using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_Exercise.Model
{
    public class MainModel
    {
        public PathModel Name { get; set; }
        public ObservableCollection<SheetModel> Sheets { get; set; }
        public SheetModel SelectSheet { get; set; }

        public MainModel(ObservableCollection<SheetModel> sheets, PathModel namepath)
        {           
            this.Sheets = sheets;
            SelectSheet = sheets.First();     
            Name = namepath;
        }
    }
}
