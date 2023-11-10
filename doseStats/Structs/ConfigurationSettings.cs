using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.Types;

namespace doseStats.Structs
{
    public struct Parameters
    {
        public string patientDataBase;
        public string documentation;
        public string excelTemplate;
        public string secondCheckTemplate;
        public bool useCurrentActivity;
        public double doseTolerance;
        public double EBRTdosePerFx;
        public int EBRTnumFx;
        //available structures to query
        public List<string> structures;
        //requested statistic lists
        public List<Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>> defaultStats;
        //aims and limits
        public List<Tuple<string, string, double, string, string, string>> aimsLimits;
        //excel configuration
        public string excelWriteFormat;
        //vectors to hold the statistics the user wants to save and eventually write to the excel file
        public List<Tuple<string, string, double, string, int, string>> excelStatistics;
        public Tuple<int, string> excelPatientName;
        public Tuple<int, string> excelPatientMRN;
        public Tuple<int, string> excelPhysician;
        public Tuple<int, string> excelDate;
        public Tuple<int, string> excelTxSummary;
        public Tuple<int, string> excelEBRTdosePerFx;
        public Tuple<int, string> excelEBRTnumFx;
        public Tuple<int, string> excelEBRTtotalDose;
        public Tuple<int, string> excelNeedleContr;
        public Tuple<int, string> excelNumNeedles;
        //list to pimrary oncologist IDs
        public Dictionary<string, string> physicianIDs;

        public void initialize()
        {
            patientDataBase = String.Format(@"\\enterprise.stanfordmed.org\depts\RadiationTherapy\Public\CancerCTR\Brachytherapy\Patient Database\Gyn Database\{0} Patient Files for ARIA\", DateTime.Now.Year.ToString());
            documentation = @"\\enterprise.stanfordmed.org\depts\RadiationTherapy\Public\Users\ESimiele\Rotation 8\ESAPI\documentation\doseStats_guide.pdf";
            excelTemplate = "DO NOT USE Gyn HDR BT current 01_2020.xls";
            secondCheckTemplate = "";
            useCurrentActivity = true;
            doseTolerance = 0.0;
            EBRTdosePerFx = 1.8;
            EBRTnumFx = 27;
            excelWriteFormat = "columnwise";
            structures = new List<string> { "bladder", "bowel", "rectum", "sigmoid", "ctv", "pt A" };
            physicianIDs = new Dictionary<string, string> { };
            defaultStats = new List<Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>>
                {
                    new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>("bladder", 3.0, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>{new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Dose at Volume (Gy)", 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute) }),
                    new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>("bowel", 3.0, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>{new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Dose at Volume (Gy)", 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute) }),
                    new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>("ctv", 10.0, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>{
                        new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Dose at Volume (Gy)", 98.0, VolumePresentation.Relative, DoseValuePresentation.Absolute),
                        new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Dose at Volume (Gy)", 90.0, VolumePresentation.Relative, DoseValuePresentation.Absolute),
                        new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Volume at Dose (%)", 100.0, VolumePresentation.Relative, DoseValuePresentation.Relative),
                        new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Volume at Dose (%)", 200.0, VolumePresentation.Relative, DoseValuePresentation.Relative),
                        new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Volume (cc)", 0.0, VolumePresentation.Relative, DoseValuePresentation.Relative) }),
                    new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>("pt A", 10.0, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>{new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Dmean (Gy)", 0.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute) }),
                    new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>("rectum", 3.0, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>{new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Dose at Volume (Gy)", 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute) }),
                };
            aimsLimits = new List<Tuple<string, string, double, string, string, string>>
                {
                    new Tuple<string, string, double, string, string, string>("bladder","Dose at Volume (Gy)", 2.0, "cc", "<80", "<95"),
                    new Tuple<string, string, double, string, string, string>("bowel","Dose at Volume (Gy)", 2.0, "cc", "<70", "<75"),
                    new Tuple<string, string, double, string, string, string>("ctv","Dose at Volume (Gy)", 98.0, "%", ">75", ""),
                    new Tuple<string, string, double, string, string, string>("ctv","Dose at Volume (Gy)", 90.0, "%", ">85", "<95"),
                    new Tuple<string, string, double, string, string, string>("rectum","Dose at Volume (Gy)", 2.0, "cc", "<65", "<75")
                };
            excelStatistics = new List<Tuple<string, string, double, string, int, string>>
                {
                    new Tuple<string, string, double, string, int, string>("bladder", "Dose at Volume (Gy)", 2.0, "cc", 12, "B"),
                    new Tuple<string, string, double, string, int, string>("bowel", "Dose at Volume (Gy)", 2.0, "cc", 15, "B"),
                    new Tuple<string, string, double, string, int, string>("ctv", "Dose at Volume (Gy)", 98.0, "%", 27, "B"),
                    new Tuple<string, string, double, string, int, string>("ctv", "Dose at Volume (Gy)", 90.0, "%", 29, "B"),
                    new Tuple<string, string, double, string, int, string>("ctv", "Volume at Dose (%)", 100.0, "%", 24, "B"),
                    new Tuple<string, string, double, string, int, string>("ctv", "Volume at Dose (%)", 200.0, "%", 25, "B"),
                    new Tuple<string, string, double, string, int, string>("ctv", "Volume (cc)", 0.0, "%", 26, "B"),
                    new Tuple<string, string, double, string, int, string>("pt A", "Dmean (Gy)", 0.0, "cc", 36, "B"),
                    new Tuple<string, string, double, string, int, string>("rectum", "Dose at Volume (Gy)", 2.0, "cc", 18, "B")
                };
            excelPatientName = new Tuple<int, string>(0, "");
            excelPatientMRN = new Tuple<int, string>(0, "");
            excelPhysician = new Tuple<int, string>(0, "");
            excelDate = new Tuple<int, string>(0, "");
            excelTxSummary = new Tuple<int, string>(0, "");
            excelEBRTdosePerFx = new Tuple<int, string>(0, "");
            excelEBRTnumFx = new Tuple<int, string>(0, "");
            excelEBRTtotalDose = new Tuple<int, string>(0, "");
            excelNeedleContr = new Tuple<int, string>(0, "");
            excelNumNeedles = new Tuple<int, string>(0, "");
        }
    }
}
