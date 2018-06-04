using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EngineeringBranchAllocation
{
    public static class CasteWiseSeatAllotment
    {
        public static AllocationOfSeatsAmongCategories GetBranchSeatsWithCatogaries(double seatsAvailableInBranch)
        {
            AllocationOfSeatsAmongCategories casteWiseSeats = new AllocationOfSeatsAmongCategories();

            casteWiseSeats.CAP = Round(GetSeats(seatsAvailableInBranch, 2), 0);
            casteWiseSeats.CAPG = Round((casteWiseSeats.CAP * 33.33) / 100, 0);
            casteWiseSeats.CAP = casteWiseSeats.CAP - casteWiseSeats.CAPG;

            seatsAvailableInBranch = seatsAvailableInBranch - (casteWiseSeats.CAP+ casteWiseSeats.CAPG);

            casteWiseSeats.PH = Round(GetSeats(seatsAvailableInBranch, 3), 0);
            casteWiseSeats.PHG = Round((casteWiseSeats.PH * 33.33) / 100, 0);
            casteWiseSeats.PH = casteWiseSeats.PH - casteWiseSeats.PHG;

            casteWiseSeats.NCC = Round(GetSeats(seatsAvailableInBranch, 1), 0);
            casteWiseSeats.NCCG = Round((casteWiseSeats.NCC * 33.33) / 100, 0);
            casteWiseSeats.NCC = casteWiseSeats.NCC - casteWiseSeats.NCCG;

            casteWiseSeats.SPORTS = Round(GetSeats(seatsAvailableInBranch, 0.5), 0);
            casteWiseSeats.SPORTSG = Round((casteWiseSeats.SPORTS * 33.33) / 100, 0);
            casteWiseSeats.SPORTS = casteWiseSeats.SPORTS - casteWiseSeats.SPORTSG;


            seatsAvailableInBranch = seatsAvailableInBranch - (casteWiseSeats.PH + casteWiseSeats.NCC + casteWiseSeats.SPORTS + casteWiseSeats.PHG + casteWiseSeats.NCCG + casteWiseSeats.SPORTSG);

            casteWiseSeats.OC = Round(GetSeats(seatsAvailableInBranch, 50), 0);
            casteWiseSeats.OCG = Round((casteWiseSeats.OC * 33.33) / 100, 0);
            casteWiseSeats.OC -= casteWiseSeats.OCG;

            casteWiseSeats.BCA = Round(GetSeats(seatsAvailableInBranch, 7), 0);
            casteWiseSeats.BCAG = Round((casteWiseSeats.BCA * 33.33) / 100, 0);
            casteWiseSeats.BCA -= casteWiseSeats.BCAG;

            casteWiseSeats.BCB = Round(GetSeats(seatsAvailableInBranch, 10), 0);
            casteWiseSeats.BCBG = Round((casteWiseSeats.BCB * 33.33) / 100, 0);
            casteWiseSeats.BCB -= casteWiseSeats.BCBG;

            casteWiseSeats.BCC = Round(GetSeats(seatsAvailableInBranch, 1), 0);
            casteWiseSeats.BCCG = Round((casteWiseSeats.BCC * 33.33) / 100, 0);
            casteWiseSeats.BCC -= casteWiseSeats.BCCG;

            casteWiseSeats.BCD = Round(GetSeats(seatsAvailableInBranch, 7), 0);
            casteWiseSeats.BCDG = Round((casteWiseSeats.BCD * 33.33) / 100, 0);
            casteWiseSeats.BCD -= casteWiseSeats.BCDG;

            casteWiseSeats.BCE = Round(GetSeats(seatsAvailableInBranch, 4), 0);
            casteWiseSeats.BCEG = Round((casteWiseSeats.BCE * 33.33) / 100, 0);
            casteWiseSeats.BCE -= casteWiseSeats.BCEG;

            casteWiseSeats.SC = Round(GetSeats(seatsAvailableInBranch, 15), 0);
            casteWiseSeats.SCG = Round((casteWiseSeats.SC * 33.33) / 100, 0);
            casteWiseSeats.SC -= casteWiseSeats.SCG;

            casteWiseSeats.ST = Round(GetSeats(seatsAvailableInBranch, 6), 0);
            casteWiseSeats.STG = Round((casteWiseSeats.ST * 33.33) / 100, 0);
            casteWiseSeats.ST -= casteWiseSeats.STG;

            return casteWiseSeats;

        }

        public static double Round(double value, int digits)
        {
            double pow = Math.Pow(10, digits);
            return Math.Truncate(value * pow + Math.Sign(value) * 0.5) / pow;
        }

        public static double GetSeats(double seatsAvailableInBranch, double percentage)
        {
            return (seatsAvailableInBranch * percentage) / 100;
        }
        
    }

    public class AllocationOfSeatsAmongCategories
    {
        public double CAP { get; set; }
        public double PH { get; set; }
        public double NCC { get; set; }
        public double SPORTS { get; set; }
        public double OC { get; set; }
        public double BCA { get; set; }
        public double BCB { get; set; }
        public double BCC { get; set; }
        public double BCD { get; set; }
        public double BCE { get; set; }
        public double SC { get; set; }
        public double ST { get; set; }
        public double CAPG { get; set; }
        public double PHG { get; set; }
        public double NCCG { get; set; }
        public double SPORTSG { get; set; }
        public double OCG { get; set; }
        public double BCAG { get; set; }
        public double BCBG { get; set; }
        public double BCCG { get; set; }
        public double BCDG { get; set; }
        public double BCEG { get; set; }
        public double SCG { get; set; }
        public double STG { get; set; }
    }
}
