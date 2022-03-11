using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        public static ScriptContext context = null;
        public Script()
        { }

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

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext c /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            context = c;
            //make sure the plans in context are brachytherapy plans, the currently open plan is a brachy plan, and the dose is calculated for the current plan
            if (c.BrachyPlansInScope.Count() == 0)
            {
                MessageBox.Show("No Brachytherapy plans in scope. Open the course that contains the brachytherapy plan, bring it into context, and try again!");
                return;
            }
            if (c.PlanSetup.GetType() != typeof(BrachyPlanSetup))
            {
                MessageBox.Show("Brachytherapy plan NOT open in window. Bring the plan of interest into context and rerun the script");
                return;
            }
            if (!c.BrachyPlanSetup.IsDoseValid)
            {
                MessageBox.Show("Dose has NOT been calculated for the open plan! Calculate the dose and try again!");
                return;
            }

            Parameters p = new Parameters();
            p.initialize();
            if (File.Exists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\configuration\\HDR_doseStats_config.ini"))
            {
                try
                {
                    using (StreamReader reader = new StreamReader(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\configuration\\HDR_doseStats_config.ini"))
                    {
                        List<Tuple<string, double, string, double, string>> sortedList = new List<Tuple<string, double, string, double, string>> { };
                        List<Tuple<string, string, double, string, int, string>> excelList_temp = new List<Tuple<string, string, double, string, int, string>> { };
                        List<Tuple<string, string, double, string, string, string>> aimsLimits_temp = new List<Tuple<string, string, double, string, string, string>> { };
                        List<string> structures_temp = new List<string> { };
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            //this line contains useful information (i.e., it is not a comment)
                            if (!String.IsNullOrEmpty(line) && line.Substring(0, 1) != "%")
                            {
                                //this line contains a path or is setting a variable
                                if (line.Contains("="))
                                {
                                    string parameter = line.Substring(0, line.IndexOf("="));
                                    string value = line = cropLine(line, "=").Trim(null);
                                    if (parameter == "patient database")
                                    {
                                        //parse patient database path
                                        p.patientDataBase = value;
                                        if (p.patientDataBase.LastIndexOf("\\") != p.patientDataBase.Length - 1) p.patientDataBase += "\\";
                                        p.patientDataBase = value;
                                        //if the path contains the string '{0}', this indicates the user intends for the code to replace that string with the current year
                                        if (p.patientDataBase.Contains("{0}")) p.patientDataBase = string.Format(p.patientDataBase, DateTime.Now.Year.ToString());
                                    }
                                    else if (parameter == "documentation path") p.documentation = value;
                                    //if it is not the database path, it is likely the excel template file name or the second check template name
                                    else if (parameter == "excel template") p.excelTemplate = value;
                                    else if (parameter == "second check template") p.secondCheckTemplate = value;
                                    else if (parameter == "use current activity for second dose calc") { if(value == "false") p.useCurrentActivity = false; }
                                    else if (parameter == "tolerance") p.doseTolerance = double.Parse(value);
                                    else if (parameter == "EBRT dose per fraction") p.EBRTdosePerFx = double.Parse(value);
                                    else if (parameter == "EBRT num fx") p.EBRTnumFx = int.Parse(value);
                                    else if (parameter.Contains("excel"))
                                    {
                                        if (parameter == "excel write format" && (value == "columnwise" || value == "rowwise")) p.excelWriteFormat = value;
                                        else
                                        {
                                            line = cropLine(line, "{");
                                            if (int.TryParse(line.Substring(0, line.IndexOf(",")), out int temp))
                                            {
                                                line = cropLine(line, ",");
                                                string temp2 = line.Substring(0, line.IndexOf("}"));
                                                if (parameter.Contains("patient name")) p.excelPatientName = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("patient MRN")) p.excelPatientMRN = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("physician")) p.excelPhysician = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("date")) p.excelDate = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("Tx summary")) p.excelTxSummary = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("EBRT dose per fx")) p.excelEBRTdosePerFx = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("EBRT num fx")) p.excelEBRTnumFx = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("EBRT total dose")) p.excelEBRTtotalDose = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("add needle contribution")) p.excelNeedleContr = new Tuple<int, string>(temp, temp2);
                                                else if (parameter.Contains("add num needles")) p.excelNumNeedles = new Tuple<int, string>(temp, temp2);
                                            }
                                            else { MessageBox.Show(String.Format("Failed to parse: '{0}' to int!", line.Substring(0, line.IndexOf(",") + 1))); }
                                        }
                                    }
                                }
                                else if(line.Contains("add structure"))
                                {
                                    line = cropLine(line, "{");
                                    structures_temp.Add(line.Substring(0, line.IndexOf("}")));
                                }
                                else if (line.Contains("add default statistic"))
                                {
                                    //add a default objective
                                    string structure;
                                    string statistic;
                                    string units;
                                    //remove everything in the line up to and including the '{' character
                                    line = cropLine(line, "{");
                                    //the requested structure should be all characters up to the first comma
                                    structure = line.Substring(0, line.IndexOf(","));
                                    //remove everything in the line up to and including the first ',' character
                                    line = cropLine(line, ",");
                                    //ensure the parse structure is one of the acceptable structures for the script. All the messagebox statements are used for error reporting and troubleshooting
                                    //if (structure.ToLower() == "bladder" || structure.ToLower() == "bowel" || structure.ToLower() == "rectum" || structure.ToLower() == "sigmoid" || structure.ToLower() == "ctv" || structure.ToLower() == "pt a")
                                    //{
                                        //try parsing the alpha/beta ratio
                                        if (double.TryParse(line.Substring(0, line.IndexOf(",")), out double ab))
                                        {
                                            //ensure the parsed alpha/beta ratio is appropriate for the parsed structure (10 for targets and 3 otherwise)
                                            //if (((structure.ToLower() == "ctv" || structure.ToLower() == "pt a") && ab == 10.0) || ((structure.ToLower() == "bladder" || structure.ToLower() == "bowel" || structure.ToLower() == "rectum" || structure.ToLower() == "sigmoid") && ab == 3.0))
                                            //{
                                                //remove everything in the line up to and including the first ',' character
                                                line = cropLine(line, ",");
                                                //grab the requested statistic
                                                statistic = line.Substring(0, line.IndexOf(","));
                                                //ensure the requested statistic is acceptable for this script
                                                if ((statistic.ToLower().Contains("volume at dose") || statistic.ToLower().Contains("dose at volume") || statistic.ToLower().Contains("dmean")) && (statistic.ToLower().Contains("(gy)") || statistic.ToLower().Contains("(%)")) || statistic.ToLower().Contains("volume (cc)"))
                                                {
                                                    //remove everything in the line up to and including the first ',' character
                                                    line = cropLine(line, ",");
                                                    //try parsing the query value
                                                    if (double.TryParse(line.Substring(0, line.IndexOf(",")), out double query_val))
                                                    {
                                                        //remove everything in the line up to and including the first ',' character
                                                        line = cropLine(line, ",");
                                                        //grab the units on the query value
                                                        units = line.Substring(0, line.IndexOf("}"));
                                                        //verify the requested units on the query value are either %, Gy, or cc.
                                                        //if the added default statistic meets all of these requirements, add it to the list of requested default statistics
                                                        if (units.ToLower().Contains("%") || units.ToLower().Contains("gy") || units.ToLower().Contains("cc")) sortedList.Add(new Tuple<string, double, string, double, string>(structure, ab, statistic, query_val, units));
                                                        else MessageBox.Show(String.Format("{0}, {1}, {2}, {3}, {4}", structure, ab, statistic, query_val, units));
                                                    }
                                                    else MessageBox.Show(String.Format("{0}, {1}, {2}, {3}", structure, ab, statistic, line.Substring(0, line.IndexOf(","))));
                                                }
                                                else MessageBox.Show(String.Format("{0}, {1}, {2}", structure, ab, statistic));
                                            //}
                                            //else MessageBox.Show(String.Format("{0}, {1}", structure, ab));
                                        }
                                        else MessageBox.Show(String.Format("{0}, {1}", structure, line.Substring(0, line.IndexOf(","))));
                                    //}
                                    //else MessageBox.Show(structure);
                                }
                                else if (line.Contains("add excel statistic"))
                                {
                                    string structure;
                                    string statistic;
                                    string units;
                                    string column;
                                    //remove everything in the line up to and including the '{' character
                                    line = cropLine(line, "{");
                                    //the requested structure should be all characters up to the first comma
                                    structure = line.Substring(0, line.IndexOf(","));
                                    //remove everything in the line up to and including the first ',' character
                                    line = cropLine(line, ",");
                                    //ensure the parse structure is one of the acceptable structures for the script. All the messagebox statements are used for error reporting and troubleshooting
                                    //if (structure.ToLower() == "bladder" || structure.ToLower() == "bowel" || structure.ToLower() == "rectum" || structure.ToLower() == "sigmoid" || structure.ToLower() == "ctv" || structure.ToLower() == "pt a")
                                    //{
                                        //grab the requested statistic
                                        statistic = line.Substring(0, line.IndexOf(","));
                                        //ensure the requested statistic is acceptable for this script
                                        if ((statistic.ToLower().Contains("volume at dose") || statistic.ToLower().Contains("dose at volume") || statistic.ToLower().Contains("dmean")) && (statistic.ToLower().Contains("(gy)") || statistic.ToLower().Contains("(%)")) || statistic.ToLower().Contains("volume (cc)"))
                                        {
                                            //remove everything in the line up to and including the first ',' character
                                            line = cropLine(line, ",");
                                            //try parsing the query value
                                            if (double.TryParse(line.Substring(0, line.IndexOf(",")), out double query_val))
                                            {
                                                //remove everything in the line up to and including the first ',' character
                                                line = cropLine(line, ",");
                                                //grab the units on the query value
                                                units = line.Substring(0, line.IndexOf(","));
                                                //verify the requested units on the query value are either %, Gy, or cc.
                                                if (units.ToLower().Contains("%") || units.ToLower().Contains("gy") || units.ToLower().Contains("cc"))
                                                {
                                                    line = cropLine(line, ",");
                                                    //mp error checking of row or column values
                                                    if (int.TryParse(line.Substring(0, line.IndexOf(",")), out int row))
                                                    {
                                                        line = cropLine(line, ",");
                                                        column = line.Substring(0, line.IndexOf("}"));
                                                        excelList_temp.Add(new Tuple<string, string, double, string, int, string>(structure, statistic, query_val, units, row, column));
                                                    }
                                                    else MessageBox.Show(String.Format("{0}, {1}, {2}, {3}, {4}", structure, statistic, query_val, units, line.Substring(0, line.IndexOf(","))));
                                                }
                                                else MessageBox.Show(String.Format("{0}, {1}, {2}, {3}", structure, statistic, query_val, units));
                                            }
                                            else MessageBox.Show(String.Format("{0}, {1}, {2}", structure, statistic, line.Substring(0, line.IndexOf(","))));
                                        }
                                        else MessageBox.Show(String.Format("{0}, {1}", structure, statistic));
                                    //}
                                    //else MessageBox.Show(structure);
                                }
                                else if (line.Contains("add limit"))
                                {
                                    string structure;
                                    string statistic;
                                    string units;
                                    string aim, limit;
                                    //remove everything in the line up to and including the '{' character
                                    line = cropLine(line, "{");
                                    //the requested structure should be all characters up to the first comma
                                    structure = line.Substring(0, line.IndexOf(","));
                                    //remove everything in the line up to and including the first ',' character
                                    line = cropLine(line, ",");
                                    //ensure the parse structure is one of the acceptable structures for the script. All the messagebox statements are used for error reporting and troubleshooting
                                    //if (structure.ToLower() == "bladder" || structure.ToLower() == "bowel" || structure.ToLower() == "rectum" || structure.ToLower() == "sigmoid" || structure.ToLower() == "ctv" || structure.ToLower() == "pt a")
                                    //{
                                        //grab the requested statistic
                                        statistic = line.Substring(0, line.IndexOf(","));
                                        //ensure the requested statistic is acceptable for this script
                                        if ((statistic.ToLower().Contains("volume at dose") || statistic.ToLower().Contains("dose at volume") || statistic.ToLower().Contains("dmean")) && (statistic.ToLower().Contains("(gy)") || statistic.ToLower().Contains("(%)")) || statistic.ToLower().Contains("volume (cc)"))
                                        {
                                            //remove everything in the line up to and including the first ',' character
                                            line = cropLine(line, ",");
                                            //try parsing the query value
                                            if (double.TryParse(line.Substring(0, line.IndexOf(",")), out double query_val))
                                            {
                                                //remove everything in the line up to and including the first ',' character
                                                line = cropLine(line, ",");
                                                //grab the units on the query value
                                                units = line.Substring(0, line.IndexOf(","));
                                                //verify the requested units on the query value are either %, Gy, or cc.
                                                if (units.ToLower().Contains("%") || units.ToLower().Contains("gy") || units.ToLower().Contains("cc"))
                                                {
                                                    //grab the aim and limits (there is no error checking on the format of these parameters)
                                                    line = cropLine(line, ",");
                                                    if (line.Substring(0, 1) != ",") aim = line.Substring(0, line.IndexOf(",")); 
                                                    else aim = "";
                                                    line = cropLine(line, ",");
                                                    if (line.Substring(0, 1) != "}") limit = line.Substring(0, line.IndexOf("}"));
                                                    else limit = "";
                                                    aimsLimits_temp.Add(new Tuple<string, string, double, string, string, string>(structure, statistic, query_val, units, aim, limit));
                                                }
                                                else MessageBox.Show(String.Format("{0}, {1}, {2}, {3}", structure, statistic, query_val, units));
                                            }
                                            else MessageBox.Show(String.Format("{0}, {1}, {2}", structure, statistic, line.Substring(0, line.IndexOf(","))));
                                        }
                                        else MessageBox.Show(String.Format("{0}, {1}", structure, statistic));
                                    //}
                                    //else MessageBox.Show(structure);
                                }
                            }
                        }
                        
                        //verify the specified patient database and excel template files exist. NOTE: THIS CODE ASSUMES THE EXCELTEMPLATE RESIDES IN THE PATIENTDATABASE DIRECTORY
                        if (!Directory.Exists(p.patientDataBase)) { MessageBox.Show("Error! Specified directory in configuration file does not exist!\nAssuming default directory!"); p.patientDataBase = ""; }
                        if (p.patientDataBase == "" || !File.Exists(Path.Combine(p.patientDataBase, p.excelTemplate))) { MessageBox.Show("Error! Specified excel template in configuration file does not exist!\nAssuming default excel template!"); p.excelTemplate = ""; }
                        if (p.documentation == "" || !File.Exists(p.documentation)) { MessageBox.Show("Warning! Specified help documentation file does not exist or can't be found!\nHelp PDF will not be accessable!"); p.documentation = ""; }

                        //If there are any items in the default statistics list, we need to sort them alphabetically by the structure id and convert the array that can be used by the script.
                        //This code is almost a direct copy of the code in the calculateStatistics_Click method of the stats class
                        if (sortedList.Any())
                        {
                            p.defaultStats.Clear();
                            //sort the list according to the structure Id's (i.e., item 1 in the sorted arrays)
                            sortedList.Sort((x, y) => x.Item1.ToLower().CompareTo(y.Item1.ToLower()));

                            //grab the first structure Id and the first alpha/beta ratio in the sorted requested statistic list
                            string temp = sortedList.First().Item1;
                            double temp2 = sortedList.First().Item2;
                            //this list is used to hold the requested statistics FOR A GIVEN STRUCTURE! A copy of this list will be added to the final requested statistics list
                            //structure id, alpha/beta, statistic requested, query value, Volume representation (relative or absolute), DoseValuePresentation
                            List<Tuple<string, double, VolumePresentation, DoseValuePresentation>> listTemp = new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>> { };
                            foreach (Tuple<string, double, string, double, string> itr in sortedList)
                            {
                                //if the current structure Id is not equal to 'temp', that means the requested statistics for the previous structure have been exhausted and we need to create a new entry in the final requested statistics
                                //array storing the static information once
                                if (itr.Item1 != temp)
                                {
                                    p.defaultStats.Add(new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>(temp, temp2, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>(listTemp)));
                                    //clear the vector of the dynamic attributes of the requested statistics. Update temp and temp2 to the current structure in the list
                                    listTemp.Clear();
                                    temp = itr.Item1;
                                    temp2 = itr.Item2;
                                }

                                //grabbing the requested statistic value is easy enough, determining the requested units on that statistic requires a bit more logic. Also need to determine the units on the query statistic
                                double stat = itr.Item4;
                                DoseValuePresentation dvp;
                                VolumePresentation vp;

                                if (itr.Item3.ToLower().Contains("dose at volume") || itr.Item3.ToLower().Contains("dmean"))
                                {
                                    //the query statistic is a volume or a mean dose
                                    //determine if the query dose is absolute or relative
                                    if (itr.Item3.ToLower().Contains("gy")) dvp = DoseValuePresentation.Absolute;
                                    else dvp = DoseValuePresentation.Relative;
                                    //determine the units of the requested volume if applicable
                                    if (itr.Item5.ToLower().Contains("cc")) vp = VolumePresentation.AbsoluteCm3;
                                    else vp = VolumePresentation.Relative;
                                }
                                else
                                {
                                    //the query statistic is a dose or the volume of the structure was requested
                                    //determine if the query volume is absolute or relative
                                    if (itr.Item3.ToLower().Contains("cc")) vp = VolumePresentation.AbsoluteCm3;
                                    else vp = VolumePresentation.Relative;
                                    //determine the units of the requested dose if applicable
                                    if (itr.Item5.ToLower().Contains("gy")) dvp = DoseValuePresentation.Absolute;
                                    else dvp = DoseValuePresentation.Relative;
                                }
                                //MessageBox.Show(string.Format("{0}, {1}, {2}, {3}, {4}, {5}", itr.Item1, itr.Item2, itr.Item3, stat, vp.ToString(), dvp.ToString()));
                                //add a new entry to the requested statistics for this specific structure
                                listTemp.Add(new Tuple<string, double, VolumePresentation, DoseValuePresentation>(itr.Item3, stat, vp, dvp));
                            }
                            //need one more add statement to ensure the final structure in the requested statistic list gets added
                            p.defaultStats.Add(new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>(temp, temp2, listTemp));
                        }
                        if(excelList_temp.Any())
                        {
                            p.excelStatistics.Clear();
                            //sort the list according to the structure Id's (i.e., item 1 in the sorted arrays)
                            excelList_temp.Sort((x, y) => x.Item1.ToLower().CompareTo(y.Item1.ToLower()));
                            p.excelStatistics = new List<Tuple<string, string, double, string, int, string>>(excelList_temp);
                        }
                        if (aimsLimits_temp.Any())
                        {
                            p.aimsLimits.Clear();
                            //sort the list according to the structure Id's (i.e., item 1 in the sorted arrays)
                            aimsLimits_temp.Sort((x, y) => x.Item1.ToLower().CompareTo(y.Item1.ToLower()));
                            p.aimsLimits = new List<Tuple<string, string, double, string, string, string>>(aimsLimits_temp);
                        }
                        if(structures_temp.Any())
                        {
                            p.structures.Clear();
                            p.structures = new List<string>(structures_temp);
                        }
                    }
                }
                catch (Exception e) { MessageBox.Show(String.Format("Error could not load configuration file because: {0}\n\nAssuming default parameters", e.Message)); }
            }
            else { MessageBox.Show("No configuration file found!\nAssuming default parameters"); }
            doseStats.stats stats = new doseStats.stats(c.BrachyPlanSetup, p);
            if(!stats.formatError) stats.ShowDialog();
            else
            {
                doseStats.doseCalc calcWindow = new doseStats.doseCalc(c.BrachyPlanSetup, p.patientDataBase, p.useCurrentActivity, p.doseTolerance, p.secondCheckTemplate);
                calcWindow.ShowDialog();
            }
        }

        private string cropLine(string line, string cropChar) { return line.Substring(line.IndexOf(cropChar) + 1, line.Length - line.IndexOf(cropChar) - 1); }

        static public ScriptContext GetScriptContext() { return context; }

    }
}