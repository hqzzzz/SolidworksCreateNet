using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace SolidNet
{
    public class SwNetData:INotifyPropertyChanged
    {
        private double _H;
        private double _Hs1;
        private double _Hs2;
        private double _Hd1;
        private double _Hd2;
        private double _D1;
        private double _D2;
        private double _L1;

        public SwNetData( )
        {
         
        }

        public void SetData(SwNetData _swNet)
        {

            H=_swNet.H;
            Hs1=_swNet.Hs1;
            Hs2=_swNet.Hs2;
            Hd1=_swNet.Hd1;
            Hd2=_swNet.Hd2;
            D1=_swNet.D1;
            D2=_swNet.D2;
            L1=_swNet.L1;
        }

   
        public double H {
            get { return _H; }
            set
            {
                _H = value;
                NotifyPropertyChanged(() => H);
            }
        }
        public double Hs1
        {
            get { return _Hs1; }
            set
            {
                _Hs1 = value;
                NotifyPropertyChanged(() => Hs1);
            }
        }
        public double Hs2
        {
            get { return _Hs2; }
            set
            {
                _Hs2 = value;
                NotifyPropertyChanged(() => Hs2);
            }
        }
        public double Hd1
        {
            get { return _Hd1; }
            set
            {
                _Hd1 = value;
                NotifyPropertyChanged(() => Hd1);
            }
        }
        public double Hd2
        {
            get { return _Hd2; }
            set
            {
                _Hd2 = value;
                NotifyPropertyChanged(() => Hd2);
            }
        }
        public double D1
        {
            get { return _D1; }
            set
            {
                _D1 = value;
                NotifyPropertyChanged(() => D1);
            }
        }
        public double D2
        {
            get { return _D2; }
            set
            {
                _D2 = value;
                NotifyPropertyChanged(() => D2);
            }
        }
        public double L1
        {
            get { return _L1; }
            set
            {
                _L1=value;
                NotifyPropertyChanged(() => L1);
            }
        }




        internal double h { get { return _H / 1000.0; } }
        internal double hs1 { get { return _Hs1 / 1000.0; } }
        internal double hs2 { get { return _Hs2 / 1000.0; } }
        internal double hd1 { get { return _Hd1; } }
        internal double hd2 { get { return _Hd2; } }
        internal double d1 { get { return _D1 /1000.0; } }
        internal double d2 { get { return _D2 / 1000.0; } }

        internal double l1 { get { return _L1 / 1000.0; } }



        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyPropertyChanged<T>(Expression<Func<T>> property)
        {
            if (PropertyChanged == null)
                return;

            var memberExpression = property.Body as MemberExpression;

            if (memberExpression == null)
                return;

            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(memberExpression.Member.Name));
        }

    }


    public class JsonClassData
    {

        public string D1 { get; set; }
        public string D2;
        public string Net;

    }
}
