﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LGchem2
{
    public class PdfDt
    {
        public string pdf_name { get; set; }
        public string pdf_path { get; set; }
        public DataTable dt_raw { get; set; }
        public DataTable dt_imp { get; set; }
        public DataTable dt_ref { get; set; }
        public double? spc { get; set; }
        public double? lcl { get; set; }
        public string ref_name { get; set; }
        public DataTable dt_absChk { get; set; }
    }

    public class Model_pdf
    {
        public string pdf_name { get; set; }
        public string pdf_path { get; set; }
    }

    public class Pgb_Val : INotifyPropertyChanged
    {
        public double _val { get; set; }
        public string _str { get; set; }
        public bool _isindertate { get; set; }

        public double val
        {
            get { return _val; }
            set { _val = value; RaisePropertyChangedEvent("val"); }
        }
        public string str
        {
            get { return _str; }
            set { _str = value; RaisePropertyChangedEvent("str"); }
        }
        public bool isindertate
        {
            get { return _isindertate; }
            set { _isindertate = value; RaisePropertyChangedEvent("isindertate"); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        void RaisePropertyChangedEvent(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
