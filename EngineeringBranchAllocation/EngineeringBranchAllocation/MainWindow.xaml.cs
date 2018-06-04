using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.IO;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Data;
using Xl = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using MahApps.Metro.Controls;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Threading;
namespace EngineeringBranchAllocation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        #region FIELDS

        public ObservableCollection<AllotmentResults> FinalAllotmentResults { get; set; }
        public ObservableCollection<StudentDetails> AllStudentDetails { get; set; }
        public List<AllocationOfSeatsAmongCategories> AllAvailableSeats { get; set; }
        public List<ObservableCollection<AllotmentResults>> AllAllotmentResults { get; set; }


        private AllocationOfSeatsAmongCategories ECEavailableSeats { get; set; }
        private AllocationOfSeatsAmongCategories CSEavailableSeats  { get; set; }
        private AllocationOfSeatsAmongCategories EEEavailableSeats  { get; set; }
        private AllocationOfSeatsAmongCategories MEavailableSeats { get; set; }
        private AllocationOfSeatsAmongCategories CEavailableSeats { get; set; }
        private AllocationOfSeatsAmongCategories ChEavailableSeats { get; set; }
        private AllocationOfSeatsAmongCategories MMEavailableSeats { get; set; }
        public ObservableCollection<AllotmentResults> EceAllotmentResults { get; set; }
        public ObservableCollection<AllotmentResults> CseAllotmentResults { get; set; }
        public ObservableCollection<AllotmentResults> EeeAllotmentResults { get; set; }
        public ObservableCollection<AllotmentResults> MeAllotmentResults { get; set; }
        public ObservableCollection<AllotmentResults> CeAllotmentResults { get; set; }
        public ObservableCollection<AllotmentResults> CheAllotmentResults { get; set; }
        public ObservableCollection<AllotmentResults> MmeAllotmentResults { get; set; }

        #endregion

        public MainWindow()
        {
            InitializeComponent();
            AllStudentDetails = new ObservableCollection<StudentDetails>();
            FinalAllotmentResults = new ObservableCollection<AllotmentResults>();
            Allotment.DataContext = FinalAllotmentResults;
            InitializeAvailableSeatsAndAllotmentResults();
            //AllAvailableSeats = new List<AllocationOfSeatsAmongCategories> { ECEavailableSeats, CSEavailableSeats, EEEavailableSeats, MEavailableSeats, CEavailableSeats, ChEavailableSeats, MMEavailableSeats };
            //AllAllotmentResults = new List<ObservableCollection<AllotmentResults>> { EceAllotmentResults, CseAllotmentResults, EeeAllotmentResults, MeAllotmentResults, CeAllotmentResults, CheAllotmentResults, MmeAllotmentResults };
        }

        private void InitializeAvailableSeatsAndAllotmentResults()
        {
            ECEavailableSeats = CasteWiseSeatAllotment.GetBranchSeatsWithCatogaries(AvailableBranchWiseSeats.ECE);
            CSEavailableSeats = CasteWiseSeatAllotment.GetBranchSeatsWithCatogaries(AvailableBranchWiseSeats.CSE);
            EEEavailableSeats = CasteWiseSeatAllotment.GetBranchSeatsWithCatogaries(AvailableBranchWiseSeats.EEE);
            MEavailableSeats = CasteWiseSeatAllotment.GetBranchSeatsWithCatogaries(AvailableBranchWiseSeats.ME);
            CEavailableSeats = CasteWiseSeatAllotment.GetBranchSeatsWithCatogaries(AvailableBranchWiseSeats.CE);
            ChEavailableSeats = CasteWiseSeatAllotment.GetBranchSeatsWithCatogaries(AvailableBranchWiseSeats.CHE);
            MMEavailableSeats = CasteWiseSeatAllotment.GetBranchSeatsWithCatogaries(AvailableBranchWiseSeats.MME);
            EceAllotmentResults = new ObservableCollection<AllotmentResults>();
            CseAllotmentResults = new ObservableCollection<AllotmentResults>();
            EeeAllotmentResults = new ObservableCollection<AllotmentResults>();
            MeAllotmentResults = new ObservableCollection<AllotmentResults>();
            CeAllotmentResults = new ObservableCollection<AllotmentResults>();
            CheAllotmentResults = new ObservableCollection<AllotmentResults>();
            MmeAllotmentResults = new ObservableCollection<AllotmentResults>();

            AllAvailableSeats = new List<AllocationOfSeatsAmongCategories> { ECEavailableSeats, CSEavailableSeats, EEEavailableSeats, MEavailableSeats, CEavailableSeats, ChEavailableSeats, MMEavailableSeats };
            AllAllotmentResults = new List<ObservableCollection<AllotmentResults>> { EceAllotmentResults, CseAllotmentResults, EeeAllotmentResults, MeAllotmentResults, CeAllotmentResults, CheAllotmentResults, MmeAllotmentResults };

        }

        private void DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex()).ToString();
        }

        private void OnButtonBranchPreferenceClick(object sender, RoutedEventArgs e)
        {
            var name = (sender as Button).Name;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";//"Text files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                if (name == btnBranchPreference.Name)
                {
                    txtblkBranchPreference.Text = openFileDialog.FileName;
                }
            }
        }

        private void OnButtonbtnGetResultsClick(object sender, System.Windows.RoutedEventArgs e)
        {
            AllStudentDetails.Clear();
            FinalAllotmentResults.Clear();
            InitializeAvailableSeatsAndAllotmentResults();
            ReadAndUpdateStudentData();
            AssignBranchToStudent();
        }

        #region AssignBranchToStudent

        private void AssignBranchToStudent()
        {
            AllocateStudentsToECE();
        }

        private void AllocateStudentsToECE()
        {
            AllocateSpecialCategorySeats();
            foreach (var availableSeats in AllAvailableSeats)
            {
                AdjustSpecialCategorySeats(availableSeats);
            }
            AllocateBranchAndUpdateCollection();
            foreach (var availableSeats in AllAvailableSeats)
            {
                if (availableSeats.OCG > 0)
                {
                    availableSeats.OC += availableSeats.OCG;
                }
            }

            //Above test code
            AllocateBranchAndUpdateCollection();
            foreach (var availableSeats in AllAvailableSeats)
            {
                AdjustAvailableSeatsAfterallotment(availableSeats);
            }
            
            AllocateBranchAndUpdateCollection();
            foreach (var availableSeats in AllAvailableSeats)
            {
                AllocateRemainingSeatsToOpenCategory(availableSeats);
            }
            
            AllocateBranchAndUpdateCollection();
            foreach (var availableSeats in AllAvailableSeats)
            {
                if (availableSeats.OCG > 0)
                {
                    availableSeats.OC += availableSeats.OCG;
                }
            }
            AllocateBranchAndUpdateCollection();


            //AllocateSpecialCategorySeats();
            //foreach (var availableSeats in AllAvailableSeats)
            //{
            //    AdjustSpecialCategorySeats(availableSeats);
            //}
           
            //AllocateBranchAndUpdateCollection();
            //foreach (var availableSeats in AllAvailableSeats)
            //{
            //    if (availableSeats.OCG > 0)
            //    {
            //        availableSeats.OC += availableSeats.OCG;
            //    }
            //}
            AllocateBranchAndUpdateCollection();

            foreach (var AllotmentresultsCollection in AllAllotmentResults)
            {
                foreach (var result in AllotmentresultsCollection)
                {
                    FinalAllotmentResults.Add(result);
                }
            }
           
        }

        private void AllocateBranchAndUpdateCollection()
        {
            var AllStudents = AllStudentDetails.Select(x => x).Where(x => x.CGPA >= 6 && !x.isBranchAllotted)
                    .OrderByDescending(x => x.CGPA).ThenByDescending(x => x.MathsAvg).ThenByDescending(x => x.PhysicsAvg).ThenByDescending(x => x.DateOfBirth);
           
            foreach (var student in AllStudents)
            {
                var preferences = student.PreferredCourses;
                foreach (var preferredCource in preferences)
                {
                    if (preferredCource.Equals(StudentDetails.Branches.ECE))
                    {
                        CheckAndAllocatebranchToStudent(student, ECEavailableSeats, EceAllotmentResults, StudentDetails.Branches.ECE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.CSE))
                    {
                        CheckAndAllocatebranchToStudent(student, CSEavailableSeats, CseAllotmentResults, StudentDetails.Branches.CSE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.EEE))
                    {
                        CheckAndAllocatebranchToStudent(student, EEEavailableSeats, EeeAllotmentResults, StudentDetails.Branches.EEE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.ME))
                    {
                        CheckAndAllocatebranchToStudent(student, MEavailableSeats, MeAllotmentResults, StudentDetails.Branches.ME);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.ChE))
                    {
                        CheckAndAllocatebranchToStudent(student, ChEavailableSeats, CheAllotmentResults, StudentDetails.Branches.ChE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.CE))
                    {
                        CheckAndAllocatebranchToStudent(student, CEavailableSeats, CeAllotmentResults, StudentDetails.Branches.CE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.MME))
                    {
                        CheckAndAllocatebranchToStudent(student, MMEavailableSeats, MmeAllotmentResults, StudentDetails.Branches.MME);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                }

            }

        }

        private void CheckAndAllocatebranchToStudent(StudentDetails student, AllocationOfSeatsAmongCategories availableSeats, ObservableCollection<AllotmentResults> EceAllotmentResults, StudentDetails.Branches branches)
        {
            //OC
            if (availableSeats.OC > 0)
            {
                EceAllotmentResults.Add(new AllotmentResults
                {
                    StudentId = student.Id,
                    Name = student.Name,
                    AllottedBranch = branches,
                    CGPA = student.CGPA.ToString(),
                    StudentCaste = student.StudentCaste,
                    StudentCategory = student.SpecialCategory
                });
                availableSeats.OC--;
                student.isBranchAllotted = true;
                return;
            }


            if (student.GenderType.Equals(StudentDetails.Gender.Female) && availableSeats.OCG > 0)
            {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.OCG--;
                    student.isBranchAllotted = true;
                return;
            }
            
            // BCA
            if(student.StudentCaste.Equals(StudentDetails.Cast.BCA))
            {
                if (availableSeats.BCA > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCA--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if(student.StudentCaste.Equals(StudentDetails.Cast.BCA) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.BCAG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCAG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            //BCB
            if(student.StudentCaste.Equals(StudentDetails.Cast.BCB))
            {
                if (availableSeats.BCB > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCB--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if(student.StudentCaste.Equals(StudentDetails.Cast.BCB) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.BCBG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCBG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            //BCC
             if(student.StudentCaste.Equals(StudentDetails.Cast.BCC))
            {
                if (availableSeats.BCC > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCC--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if(student.StudentCaste.Equals(StudentDetails.Cast.BCC) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.BCCG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCCG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            //BCD
             if(student.StudentCaste.Equals(StudentDetails.Cast.BCD))
            {
                if (availableSeats.BCD > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCD--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if(student.StudentCaste.Equals(StudentDetails.Cast.BCD) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.BCDG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCDG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            //BCE
             if(student.StudentCaste.Equals(StudentDetails.Cast.BCE))
            {
                if (availableSeats.BCE > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCE--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if(student.StudentCaste.Equals(StudentDetails.Cast.BCE) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.BCEG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.BCEG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }

            //SC
            if(student.StudentCaste.Equals(StudentDetails.Cast.SC))
            {
                if (availableSeats.SC > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.SC--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if(student.StudentCaste.Equals(StudentDetails.Cast.SC) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.SCG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.SCG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            //ST
            if(student.StudentCaste.Equals(StudentDetails.Cast.ST))
            {
                if (availableSeats.ST > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.ST--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
           if(student.StudentCaste.Equals(StudentDetails.Cast.ST) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.STG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.STG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
        }

        private void AllocateSpecialCategorySeats()
        {
            var AllStudents = AllStudentDetails.Select(x => x).Where(x => x.CGPA >= 6 && !x.isBranchAllotted)
                   .OrderByDescending(x => x.CGPA).ThenByDescending(x => x.MathsAvg).ThenByDescending(x => x.PhysicsAvg).ThenByDescending(x => x.DateOfBirth);
            var studentsWithSpecialCatogory = AllStudents.Select(x => x).Where(x => x.SpecialCategory != StudentDetails.Category.NONE)
                  .OrderByDescending(x => x.CGPA).ThenByDescending(x => x.MathsAvg).ThenByDescending(x => x.PhysicsAvg).ThenByDescending(x => x.DateOfBirth);
            foreach (var student in studentsWithSpecialCatogory)
            {
                var preferences = student.PreferredCourses;
                foreach (var preferredCource in preferences)
                {
                    if (preferredCource.Equals(StudentDetails.Branches.ECE))
                    {
                        CheckAndAllocatebranchToSpecialCategoryStudent(student, ECEavailableSeats, EceAllotmentResults, StudentDetails.Branches.ECE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.CSE))
                    {
                        CheckAndAllocatebranchToSpecialCategoryStudent(student, CSEavailableSeats, CseAllotmentResults, StudentDetails.Branches.CSE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.EEE))
                    {
                        CheckAndAllocatebranchToSpecialCategoryStudent(student, EEEavailableSeats, EeeAllotmentResults, StudentDetails.Branches.EEE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.ME))
                    {
                        CheckAndAllocatebranchToSpecialCategoryStudent(student, MEavailableSeats, MeAllotmentResults, StudentDetails.Branches.ME);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.ChE))
                    {
                        CheckAndAllocatebranchToSpecialCategoryStudent(student, ChEavailableSeats, CheAllotmentResults, StudentDetails.Branches.ChE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.CE))
                    {
                        CheckAndAllocatebranchToSpecialCategoryStudent(student, CEavailableSeats, CeAllotmentResults, StudentDetails.Branches.CE);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                    if (preferredCource.Equals(StudentDetails.Branches.MME))
                    {
                        CheckAndAllocatebranchToSpecialCategoryStudent(student, MMEavailableSeats, MmeAllotmentResults, StudentDetails.Branches.MME);
                        if (student.isBranchAllotted)
                        {
                            break;
                        }
                    }
                }
            }
        }

        private void CheckAndAllocatebranchToSpecialCategoryStudent(StudentDetails student, AllocationOfSeatsAmongCategories availableSeats, ObservableCollection<EngineeringBranchAllocation.AllotmentResults> EceAllotmentResults, StudentDetails.Branches branches)
        {
            // CAP
            if(student.SpecialCategory.Equals(StudentDetails.Category.CAP))
            {
                if (availableSeats.CAP > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.CAP--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if(student.SpecialCategory.Equals(StudentDetails.Category.CAP) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.CAPG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.CAPG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            //PH
            if (student.SpecialCategory.Equals(StudentDetails.Category.PH) || student.SpecialCategory.Equals(StudentDetails.Category.PHHEARING)
                     || student.SpecialCategory.Equals(StudentDetails.Category.PHORTHO)
                      || student.SpecialCategory.Equals(StudentDetails.Category.PHVISUAL))
            {
                if (availableSeats.PH > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.PH--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if (student.SpecialCategory.Equals(StudentDetails.Category.PH) || student.SpecialCategory.Equals(StudentDetails.Category.PHHEARING)
                     || student.SpecialCategory.Equals(StudentDetails.Category.PHORTHO)
                      || student.SpecialCategory.Equals(StudentDetails.Category.PHVISUAL) && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.PH > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.PH--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            //NCC
            if(student.SpecialCategory == StudentDetails.Category.NCC)
            {
                if (availableSeats.NCC > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.NCC--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if(student.SpecialCategory == StudentDetails.Category.NCC && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.NCCG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.NCCG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            //Sports
            if (student.SpecialCategory == StudentDetails.Category.SPORTS)
            {
                if (availableSeats.SPORTS > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.SPORTS--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
            if (student.SpecialCategory == StudentDetails.Category.SPORTS && student.GenderType.Equals(StudentDetails.Gender.Female))
            {
                if (availableSeats.SPORTSG > 0)
                {
                    EceAllotmentResults.Add(new AllotmentResults
                    {
                        StudentId = student.Id,
                        Name = student.Name,
                        AllottedBranch = branches,
                        CGPA = student.CGPA.ToString(),
                        StudentCaste = student.StudentCaste,
                        StudentCategory = student.SpecialCategory
                    });
                    availableSeats.SPORTSG--;
                    student.isBranchAllotted = true;
                    return;
                }
            }
        }

        #endregion

        #region AdjustSeats

        private void AllocateRemainingSeatsToOpenCategory(AllocationOfSeatsAmongCategories availableSeats)
        {
            //Male
            if (availableSeats.BCA > 0)
            {
                availableSeats.OC += availableSeats.BCA;
            }
            if (availableSeats.BCB > 0)
            {
                availableSeats.OC += availableSeats.BCB;
            }
            if (availableSeats.BCC > 0)
            {
                availableSeats.OC += availableSeats.BCC;
            }
            if (availableSeats.BCD > 0)
            {
                availableSeats.OC += availableSeats.BCD;
            }
            if (availableSeats.BCE > 0)
            {
                availableSeats.OC += availableSeats.BCE;
            }
            if (availableSeats.SC > 0)
            {
                availableSeats.OC += availableSeats.SC;
            }
            if (availableSeats.ST > 0)
            {
                availableSeats.OC += availableSeats.ST;
            }


            //Female

            if (availableSeats.BCAG > 0)
            {
                availableSeats.OCG += availableSeats.BCAG;
            }
            if (availableSeats.BCBG > 0)
            {
                availableSeats.OCG += availableSeats.BCBG;
            }
            if (availableSeats.BCCG > 0)
            {
                availableSeats.OCG += availableSeats.BCCG;
            }
            if (availableSeats.BCDG > 0)
            {
                availableSeats.OCG += availableSeats.BCDG;
            }
            if (availableSeats.BCEG > 0)
            {
                availableSeats.OCG += availableSeats.BCEG;
            }
            if (availableSeats.SCG > 0)
            {
                availableSeats.OCG += availableSeats.SCG;
            }
            if (availableSeats.SCG > 0)
            {
                availableSeats.OCG += availableSeats.STG;
            }
        }

        private void AdjustAvailableSeatsAfterallotment(AllocationOfSeatsAmongCategories availableSeats)
        {
            //Male
            if (availableSeats.BCA > 0)
            {
                if (availableSeats.BCB > 0)
                {
                    if (availableSeats.BCC > 0)
                    {
                        if (availableSeats.BCD > 0)
                        {
                            if (availableSeats.BCE > 0)
                            {
                                availableSeats.BCA += availableSeats.BCE;
                            }
                            else
                            {
                                availableSeats.BCE += (availableSeats.BCA + availableSeats.BCB + availableSeats.BCC + availableSeats.BCE);
                            }
                        }
                        else
                        {
                            availableSeats.BCD += (availableSeats.BCA + availableSeats.BCB + availableSeats.BCC);
                        }
                    }
                    else
                    {
                        availableSeats.BCC += (availableSeats.BCA + availableSeats.BCB);
                    }
                }
                else
                {
                    availableSeats.BCB += availableSeats.BCA;
                }
            }
            if (availableSeats.SC > 0)
            {
                availableSeats.ST += availableSeats.SC;
            }
            else if (availableSeats.SC > 0)
            {
                availableSeats.SC += availableSeats.ST;
            }




            if (availableSeats.BCAG > 0)
            {
                if (availableSeats.BCBG > 0)
                {
                    if (availableSeats.BCCG > 0)
                    {
                        if (availableSeats.BCDG > 0)
                        {
                            if (availableSeats.BCEG > 0)
                            {
                                availableSeats.BCAG += availableSeats.BCEG;
                            }
                            else
                            {
                                availableSeats.BCEG += (availableSeats.BCAG + availableSeats.BCBG + availableSeats.BCCG + availableSeats.BCEG);
                            }
                        }
                        else
                        {
                            availableSeats.BCDG += (availableSeats.BCAG + availableSeats.BCBG + availableSeats.BCCG);
                        }
                    }
                    else
                    {
                        availableSeats.BCCG += (availableSeats.BCAG + availableSeats.BCBG);
                    }
                }
                else
                {
                    availableSeats.BCBG += availableSeats.BCAG;
                }
            }
            if (availableSeats.SCG > 0)
            {
                availableSeats.STG += availableSeats.SCG;
            }
            else if (availableSeats.SCG > 0)
            {
                availableSeats.SCG += availableSeats.STG;
            }
        }

        private void AdjustSpecialCategorySeats(AllocationOfSeatsAmongCategories availableSeats)
        {
            //Male
            if (availableSeats.CAP > 0)
            {
                availableSeats.OC += availableSeats.CAP;
            }
            if (availableSeats.PH > 0)
            {
                availableSeats.OC += availableSeats.PH;
            }
            if (availableSeats.NCC > 0)
            {
                availableSeats.OC += availableSeats.NCC;
            }
            if (availableSeats.SPORTS > 0)
            {
                availableSeats.OC += availableSeats.SPORTS;
            }
            //Female
            if (availableSeats.CAPG > 0)
            {
                availableSeats.OCG += availableSeats.CAPG;
            }
            if (availableSeats.PHG > 0)
            {
                availableSeats.OCG += availableSeats.PHG;
            }
            if (availableSeats.NCCG > 0)
            {
                availableSeats.OCG += availableSeats.NCCG;
            }
            if (availableSeats.SPORTSG > 0)
            {
                availableSeats.OCG += availableSeats.SPORTSG;
            }
        }

        #endregion

        #region PrepareData

        private void ReadAndUpdateStudentData()
        {
            DataSet dataSet = new DataSet();
            FileStream fs = File.Open(txtblkBranchPreference.Text, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
            dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            foreach (DataTable item in dataSet.Tables)
            {
                if (item.TableName == "BRANTCH_PREFERENCES")
                {
                    UpdateStudentPreferences(item);
                }
                else
                {
                    UpdateOtherStudentDetails(item);
                }
            }
            fs.Close();
        }

        private void ObButtonbtnSaveResultsClick(object sender, System.Windows.RoutedEventArgs e)
        {
            
            //Xl._Application xlApp = new Xl.Application();
            //xlApp.Visible = false;
            //xlApp.DisplayAlerts = false;
            //Microsoft.Office.Interop.Excel.Range celLrangE; 
            //Xl.Workbook xlWorkBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Xl.Worksheet xlWorkSheet = (Xl.Worksheet)xlWorkBook.ActiveSheet;
            //xlWorkSheet.Name = "BranchAllotmentResults";
            //int rowcount = Allotment.Items.Count;

            //foreach (AllotmentResults datarow in Allotment.Items)
            //{
            //    rowcount += 1;
            //    for (int i = 1; i <= Allotment.Columns.Count; i++)
            //    {

            //        if (rowcount == 3)
            //        {
            //            xlWorkSheet.Cells[2, i] = Allotment.Columns[i - 1].Header;
            //            xlWorkSheet.Cells.Font.Color = System.Drawing.Color.Black;

            //        }

            //        xlWorkSheet.Cells[rowcount, i] = datarow[i - 1].ToString();

            //        if (rowcount > 3)
            //        {
            //            if (i == Allotment.Columns.Count)
            //            {
            //                if (rowcount % 2 == 0)
            //                {
            //                    celLrangE = xlWorkSheet.Range[xlWorkSheet.Cells[rowcount, 1], xlWorkSheet.Cells[rowcount, Allotment.Columns.Count]];
            //                }

            //            }
            //        }

            //    }

            //}


            //celLrangE = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[rowcount, Allotment.Columns.Count]];
            //celLrangE.EntireColumn.AutoFit();
            //Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
            //border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            //border.Weight = 2d;

            //celLrangE = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[2, Allotment.Columns.Count]];

            //xlWorkBook.SaveAs("C://Users/ARUN/Desktop/Anwesh/MyFinalData/anweshresult.xlsx"); ;
            //xlWorkBook.Close();
            //xlApp.Quit();

            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);

            //MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.xlsx");

          
        }

        private void UpdateStudentPreferences(DataTable item)
        {
            DataTable dt = new DataTable();
            dt = item;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                StudentDetails stuDetails;
                var studentid = dt.Rows[i][1].ToString();
                if (AllStudentDetails.Any(x => x.Id == studentid))
                {
                    stuDetails = AllStudentDetails.Select(x => x).Where(x => x.Id == studentid).FirstOrDefault();
                }
                else
                {
                    stuDetails = new StudentDetails();
                }

                for (int j = 1; j < dt.Columns.Count; j++)
                {
                    if (j == 1)
                    {
                        if (stuDetails.Id == null)
                        {
                            stuDetails.Id = dt.Rows[i][j].ToString();
                        }
                    }
                    else
                    {
                        var val = dt.Rows[i][j].ToString();
                        if (stuDetails.PreferredCourses == null)
                        {
                            stuDetails.PreferredCourses = new List<StudentDetails.Branches>();
                        }
                        stuDetails.PreferredCourses.Add((StudentDetails.Branches)Enum.Parse(typeof(StudentDetails.Branches), val));
                    }
                }

                if (!AllStudentDetails.Any(x => x.Id == studentid))
                {
                    AllStudentDetails.Add(stuDetails);
                }
            }
        }

        private void UpdateOtherStudentDetails(DataTable dt)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                StudentDetails stuDetails;
                var studentid = dt.Rows[i][1].ToString();
                if (AllStudentDetails.Any(x => x.Id == studentid))
                {
                    stuDetails = AllStudentDetails.Select(x => x).Where(x => x.Id == studentid).FirstOrDefault();
                }
                else
                {
                    stuDetails = new StudentDetails();
                }


                if (stuDetails.Id == null)
                {
                    stuDetails.Id = dt.Rows[i][1].ToString();
                    stuDetails.Name = dt.Rows[i][2].ToString();
                    stuDetails.GenderType = (StudentDetails.Gender)Enum.Parse(typeof(StudentDetails.Gender), dt.Rows[i][3].ToString());
                    stuDetails.StudentCaste = (StudentDetails.Cast)Enum.Parse(typeof(StudentDetails.Cast), dt.Rows[i][4].ToString().ToUpper().Replace("-", string.Empty));
                    stuDetails.SpecialCategory = (StudentDetails.Category)Enum.Parse(typeof(StudentDetails.Category),
                        dt.Rows[i][5].ToString() == string.Empty ? "NONE" : dt.Rows[i][5].ToString().ToUpper().Replace("-", string.Empty) );
                    stuDetails.CGPA = Convert.ToDouble(dt.Rows[i][6]);
                    if(dt.Rows[i][7].ToString() != string.Empty )
                    {
                        stuDetails.MathsAvg = (double)Convert.ToDouble(dt.Rows[i][7]);
                    }
                    if(dt.Rows[i][8].ToString() != string.Empty)
                    {
                        stuDetails.PhysicsAvg = (double)Convert.ToDouble(dt.Rows[i][8]);
                    }
                    if(dt.Rows[i][9].ToString() != string.Empty )
                    {
                        stuDetails.DateOfBirth = Convert.ToDateTime(dt.Rows[i][9]);
                    }
                }

                if (!AllStudentDetails.Any(x => x.Id == studentid))
                {
                    AllStudentDetails.Add(stuDetails);
                }
            }
        }

        #endregion
    }
}
