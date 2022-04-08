using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_Exercise.Model
{
    public class PathModel
    {
        public string Namepath { get; set; }
        public PathModel(string namepath) 
        { 
            Namepath = namepath; 
        }
    }
}
