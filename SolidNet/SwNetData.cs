using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidNet
{
    public class SwNetData
    {


        public SwNetData()
        {
         
        }

        public double H { get; set; }
        public double Hs1 { get; set; }
        public double Hs2 { get; set; }
        public double Hd1 { get; set; }
        public double Hd2 { get; set; }
        public double D1 { get; set; }
        public double D2 { get; set; }
        public double L1 { get; set; }




        public double h { get { return H / 1000.0; } }
        public double hs1 { get { return Hs1 / 1000.0; } }
        public double hs2 { get { return Hs2 / 1000.0; } }
        public double hd1 { get { return Hd1; } }
        public double hd2 { get { return Hd2; } }
        public double d1 { get { return D1 /1000.0; } }
        public double d2 { get { return D2 / 1000.0; } }

        public double l1 { get { return L1 / 1000.0; } }

    }
}
