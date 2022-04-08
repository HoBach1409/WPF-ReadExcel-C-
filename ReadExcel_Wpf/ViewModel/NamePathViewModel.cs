using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WPF_Exercise.Model;

namespace WPF_ReadExcel.ViewModel
{
    public class NamePathViewModel : Base
    {
        PathModel Model { get; set; }
        public string Namepath
        {
            get
            {
                return Model.Namepath;
            }
            set
            {
                Model.Namepath = value;
                OnPropertyChanged(nameof(Namepath));
            }
        }
        public NamePathViewModel (PathModel namepath)
        {
            Model = namepath;    
        }
    }
}


