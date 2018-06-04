using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EngineeringBranchAllocation
{
    public class AllotmentResults
    {
        public string StudentId { get; set; }

        public string Name { get; set; }

        public string CGPA { get; set; }

        public StudentDetails.Branches AllottedBranch { get; set; }

        public StudentDetails.Cast StudentCaste { get; set; }

        public StudentDetails.Category StudentCategory { get; set; }

    }
}
