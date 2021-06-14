using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JurisUtilityBase
{
    public class Row
    {
        public Row()
        {
            client = "";
            clisys = 0;
            oldTask  = "";
            newTask  = "";
            error = "";
        }
        public string client { get; set; }
        public string oldTask { get; set; }
        public string newTask { get; set; }
        public int clisys { get; set; }
        public string error { get; set; }
    }
}
