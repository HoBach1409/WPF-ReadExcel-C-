using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_Exercise.Model
{
    public class SheetModel
    {
        public string Namesheet 
            {get;set;}
        public DataTable DataSheet 
        { get;
            set;
        }
        public SheetModel(string namesheet, DataTable datasheet)
        {
            Namesheet = namesheet;
            DataSheet = datasheet;           
        }
    }
}
