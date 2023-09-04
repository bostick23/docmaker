using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocMaker.Models
{
    public class Settings
    {
        public int Radius
        {
            get; set;
        }
        public System.Drawing.Color BackgroundColor
        {
            get; set;
        }
        public System.Drawing.Color BorderColor
        {
            get; set;
        }
        public double BorderWidth
        {
            get; set;
        }
    }
}
