using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Extensions.Models.jsTree
{
    public class JSONalternativeFormat
    {
        public string id { get; set; }
        public string parent { get; set; }
        public string text { get; set; }
        public string icon { get; set; }
        public State state { get; set; }
        public bool children { get { return false; } private set { ;} }
    }

    public  class State
    {
       bool opened { get; set;}
       bool disabled { get; set;}
       bool selected { get; set; }
 
    }
}