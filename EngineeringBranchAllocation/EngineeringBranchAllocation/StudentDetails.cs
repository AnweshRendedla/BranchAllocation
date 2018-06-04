using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EngineeringBranchAllocation
{   
    public class StudentDetails
    {
        public enum Branches
        {
            ECE,
            CSE,
            EEE,
            ME,
            ChE,
            CE,
            MME
        }

        public enum Cast
        {
            OC,
            BCA,
            BCB,
            BCC,
            BCD,
            BCE,
            SC,
            ST
            //BCAD,
            //BCABD
        };
        public enum Category
        {
            NONE,
            CAP,
            PH,
            NCC,
            SPORTS,
            PHORTHO,
            PHHEARING,
            PHVISUAL,
            OH,
            VI
        }
        public enum Gender{
            Male,
            Female
        }
        public string Id { get; set; }
        public string Name { get; set; }
        public Gender GenderType { get; set; }
        public Cast StudentCaste { get; set; }
        public Category SpecialCategory { get; set; }
        public Double CGPA { get; set; }       
        public double? MathsAvg { get; set; }
        public double? PhysicsAvg { get; set; }
        public DateTime? DateOfBirth { get; set; }        
        public List<Branches> PreferredCourses { get; set; }
        public bool isBranchAllotted { get; set; }
    }
}
