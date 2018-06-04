using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EngineeringBranchAllocation
{
    public static class AvailableBranchWiseSeats
    {
        public static int TotalNumberOfSeats = 1000;

        public static double ECE
        {
            get
            {
                return 21.4 * TotalNumberOfSeats / 100;
            }
        }
        public static double EEE
        {
            get
            {
                return 14.4 * TotalNumberOfSeats / 100;
            }
        }
        public static double CSE
        {
            get
            {
                return 21.4 * TotalNumberOfSeats / 100;
            }
        }
        public static double CE
        {
            get
            {
                return 14.3 * TotalNumberOfSeats / 100;
            }
        }
        public static double CHE
        {
            get
            {
                return 7.1 * TotalNumberOfSeats / 100;
            }
        }
        public static double ME
        {
            get
            {
                return 14.3 * TotalNumberOfSeats / 100;
            }
        }
        public static double MME
        {
            get
            {
                return 7.1 * TotalNumberOfSeats / 100;
            }
        }
    }
}
