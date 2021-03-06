﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beadando_projekt
{
    public class Fire
    {
        private string _month;
        public string Name { get; set; }
        public string Cause { get; set; }
        public string Month
        {
            get { return _month; }

            set
            {
                _month = Month;

                switch (value)
                {
                    case "January":
                        _month = "Január";
                        break;

                    case "February":
                        _month = "Február";
                        break;

                    case "March":
                        _month = "Március";
                        break;

                    case "April":
                        _month = "Április";
                        break;

                    case "May":
                        _month = "Május";
                        break;

                    case "June":
                        _month = "Június";
                        break;

                    case "July":
                        _month = "Július";
                        break;

                    case "August":
                        _month = "Augusztus";
                        break;

                    case "September":
                        _month = "Szeptember";
                        break;

                    case "October":
                        _month = "Október";
                        break;

                    case "November":
                        _month = "November";
                        break;

                    case "December":
                        _month = "December";
                        break;
                    default:
                        _month = value;
                        break;
                };

            }
        }
        public int  Year { get; set; }
        public string Country { get; set; }
        public int Acres { get; set; }
        public int Structures { get; set; }
        public int Deaths { get; set; }
    }
}
