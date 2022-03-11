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
using System.Windows.Shapes;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace doseStats
{
    /// <summary>
    /// Interaction logic for doseCalc.xaml
    /// </summary>
    public partial class doseCalc : System.Windows.Window
    {
        BrachyPlanSetup plan = null;
        public string secondCheckFile = "";
        bool useCurrentActivity = true;
        string patientDataBase = "";
        public string filePath = "";
        double doseAgreeTolerance = 0.0;

        double plannedActivity = 0.0;
        double currentActivity = 0.0;
        VVector QAptLoc;
        double QAptDose = 0.0;
        double totalPlannedTime = 0.0;

        RoutedCommand writeExcelMacro = new RoutedCommand();
        RoutedCommand writeTextMacro = new RoutedCommand();
        RoutedCommand closeWindowMacro = new RoutedCommand();
        public doseCalc(BrachyPlanSetup p, string patDataBase, bool currentOrInitial, double tol, string secCheckFile)
        {
            plan = p;
            patientDataBase = patDataBase;
            secondCheckFile = secCheckFile;
            useCurrentActivity = currentOrInitial;
            doseAgreeTolerance = tol;
            plan.DoseValuePresentation = DoseValuePresentation.Absolute;
            InitializeComponent();

            writeExcelMacro.InputGestures.Add(new KeyGesture(Key.W, ModifierKeys.Control));
            writeTextMacro.InputGestures.Add(new KeyGesture(Key.T, ModifierKeys.Control));
            closeWindowMacro.InputGestures.Add(new KeyGesture(Key.Q, ModifierKeys.Control));

            CommandBindings.Add(new CommandBinding(writeExcelMacro, writeExcel_Click));
            CommandBindings.Add(new CommandBinding(writeTextMacro, writeText_Click));
            CommandBindings.Add(new CommandBinding(closeWindowMacro, closeWindow));

            if (runSecondDoseCalculation()) this.Close();
        }

        //close the dose calculation window. Called from the Ctrl + Q keyboard macro
        private void closeWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private bool runSecondDoseCalculation()
        {
            RadioactiveSource source = plan.Catheters.First().TreatmentUnit.GetActiveRadioactiveSource();
            if (source == null) { MessageBox.Show("No radioactive source present in catheter!!"); return true; }

            //get the calibration date of the source, today's date, and the interval between the two
            DateTime? calDate = source.CalibrationDate;
            DateTime today = DateTime.Now;
            TimeSpan interval;
            //get the interval between the source calibration date and today. Set to zero if calDate dose NOT have a value
            if (calDate.HasValue) interval = today.Subtract((DateTime)calDate);
            else interval = today.Subtract(today);

            string message = "";
            message += " " + DateTime.Now.ToString() + System.Environment.NewLine;
            message += String.Format(" Patient: {0}", VMS.TPS.Script.GetScriptContext().Patient.Name) + System.Environment.NewLine;
            message += String.Format(" Plan: {0}", plan.Id) + System.Environment.NewLine;
            message += System.Environment.NewLine;
            //planned activity in mCi
            plannedActivity = (source.Strength / source.RadioactiveSourceModel.ActivityConversionFactor);
            //calculate current source activity. NOTE: THE ACITIVTY CALCULATION IS PER DAY AT 12:00 AM! This matches the decay calculation in Eclipse. interval.days is an integer. Must be cast to a double
            currentActivity = (source.Strength / source.RadioactiveSourceModel.ActivityConversionFactor) * Math.Exp(-Math.Log(2) / source.RadioactiveSourceModel.HalfLife * (double)(interval.Days * 24 * 3600));
            message += String.Format(" Ir-192 planned activity: {0:0.00} Ci", plannedActivity / 1000) + System.Environment.NewLine;
            message += String.Format(" Ir-192 current activity: {0:0.00} Ci", currentActivity / 1000) + System.Environment.NewLine;
            message += System.Environment.NewLine;
            message += String.Format(" Use current activity for second dose calculation check: {0}", useCurrentActivity) + System.Environment.NewLine;
            message += System.Environment.NewLine;

            //logic to grab all reference points in a plan. There should be at most two reference lines (pt A and pt B)
            List<Structure> referenceLines = plan.StructureSet.Structures.Where(x => x.GetReferenceLinePoints().Length > 0).ToList();
            //vector to hold the point Id, location, dose, and dose difference from the prescription
            //structure id, x, y, z, dose, dose diff from prescription^2
            ReferencePoint referencePt = null;
            if (referenceLines.Count > 0)
            {
                List<Tuple<string, double, double, double, double, double>> QA = new List<Tuple<string, double, double, double, double, double>> { };
                //iterate through each reference line and add each point to the vector
                foreach (Structure s in referenceLines)
                {
                    VVector[] points = s.GetReferenceLinePoints();
                    for (int i = 0; i < points.Length; i++)
                    {
                        VVector loc = points.ElementAt(i);
                        QA.Add(new Tuple<string, double, double, double, double, double>(s.Id, loc.x, loc.y, loc.z, plan.Dose.GetDoseToPoint(loc).Dose, Math.Pow(plan.Dose.GetDoseToPoint(loc).Dose - plan.TotalDose.Dose, 2)));
                    }
                }
                //sort the vector of reference points by the square dose difference between the dose to the point and the prescription point (i.e., we want to use the point that has the closest dose to the prescription for QA calculations).
                //The point that will be used for QA should be the first item in the QA vector
                QA.Sort((x, y) => x.Item6.CompareTo(y.Item6));

                message += String.Format(" QA reference pt: {0}", referenceLines.First().Id) + System.Environment.NewLine;
                QAptLoc = new VVector(QA.First().Item2, QA.First().Item3, QA.First().Item4);
                QAptLoc = plan.StructureSet.Image.DicomToUser(QAptLoc, plan);
                message += String.Format(" QA reference location (x,y,z): ({0:0.00} cm, {1:0.00} cm, {2:0.00} cm)", QAptLoc.x / 10, QAptLoc.y / 10, QAptLoc.z / 10) + System.Environment.NewLine;
                message += System.Environment.NewLine;
            }
            else if (plan.ReferencePoints.Where(x => x.Id.ToLower().Contains("qa") || x.Id.ToLower().Contains("radcalc")).Any())
            {
                referencePt = plan.ReferencePoints.First(x => x.Id.ToLower().Contains("qa") || x.Id.ToLower().Contains("radcalc"));
                if (referencePt.HasLocation(plan))
                {
                    //write the dose to the QA point to the spreadsheet
                    message += String.Format(" QA reference pt: {0}", referencePt.Id) + System.Environment.NewLine;
                    QAptLoc = referencePt.GetReferencePointLocation(plan);
                    QAptLoc = plan.StructureSet.Image.DicomToUser(QAptLoc, plan);
                    message += String.Format(" QA reference location (x,y,z): ({0:0.00} cm, {1:0.00} cm, {2:0.00} cm)", QAptLoc.x / 10, QAptLoc.y / 10, QAptLoc.z / 10) + System.Environment.NewLine;
                    message += System.Environment.NewLine;
                }
                else MessageBox.Show(" Found QA reference point, but it has no location!");
            }
            else MessageBox.Show(" No reference lines or points found!"); 

            message += String.Format(" Ir-192 Γ-constant: {0} R-cm2/mCi-hr", 4.69) + System.Environment.NewLine;
            message += String.Format(" R->cGy in air conversion: {0} cGy/R", 0.876) + System.Environment.NewLine;
            message += String.Format(" cGy in air->cGy in tissue conversion: {0} cGy/cGy", 1.1) + System.Environment.NewLine;
            double constant = 4.69 * 0.876 * 1.1 / 3600;
            message += String.Format(" Ir-192 dose-rate constant: {0:0.00000} cGy-cm2/mCi-sec", constant) + System.Environment.NewLine;
            message += System.Environment.NewLine;
            message += String.Format(" NOTE: THIS CALCULATION USES THE POINT-SOURCE APPROXIMATION!") + System.Environment.NewLine;

            double sumDose = 0.0;
            double dose;

            foreach (Catheter c in plan.Catheters.ToList())
            {
                message += System.Environment.NewLine;
                message += String.Format(" Catheter #{0}", c.ChannelNumber) + System.Environment.NewLine;
                message += String.Format("--------------------------------------------------------------------------------------------------------") + System.Environment.NewLine;
                message += String.Format(" | Position (cm) | Planned Dt (sec) | x (cm) | y (cm) | z (cm) | r (cm) | Dose (cGy) | Tx day Dt (sec) |") + System.Environment.NewLine;
                List<SourcePosition> sourcePos = c.SourcePositions.ToList();
                foreach (SourcePosition s in sourcePos)
                {
                    //the position of the center of the source (in user coordinates!)
                    VVector pos = s.Translation;
                    pos = plan.StructureSet.Image.DicomToUser(pos, plan);
                    VVector delta = new VVector(Math.Pow((QAptLoc.x - pos.x) / 10, 2), Math.Pow((QAptLoc.y - pos.y) / 10, 2), Math.Pow((QAptLoc.z - pos.z) / 10, 2));
                    double r = Math.Sqrt(delta.x + delta.y + delta.z);
                    dose = constant * s.DwellTime / Math.Pow(r, 2);
                    if (useCurrentActivity) dose *= currentActivity;
                    else dose *= plannedActivity;
                    sumDose += dose;
                    message += String.Format(" | {0,-13:N2} | {1,-16:N1} | {2,-6:N2} | {3,-6:N2} | {4,-6:N2} | {5,-6:N2} | {6,-10:N2} | {7,-15:N1} |",
                    (c.ApplicatorLength - c.GetSourcePosCenterDistanceFromTip(s)) / 10, s.DwellTime, pos.x / 10, pos.y / 10, pos.z / 10, r, dose, s.DwellTime*plannedActivity/currentActivity) + System.Environment.NewLine;
                }
                totalPlannedTime += c.GetTotalDwellTime();
            }
            message += System.Environment.NewLine;
            message += " Summary" + System.Environment.NewLine;

            message += String.Format( "--------------------------------------------------------------------------------------------------------") + System.Environment.NewLine;
            message += String.Format(" Planned treatment time: {0:0.00} seconds", totalPlannedTime) + System.Environment.NewLine;
            message += String.Format(" Treatment time corrected for decay: {0:0.00} seconds", totalPlannedTime*plannedActivity/currentActivity) + System.Environment.NewLine;

            QAptDose = plan.Dose.GetDoseToPoint(plan.StructureSet.Image.UserToDicom(QAptLoc, plan)).Dose / (double)plan.NumberOfFractions;
            message += String.Format(" Dose at Reference Point: {0:0.00} cGy", QAptDose) + System.Environment.NewLine;
            if(useCurrentActivity) message += String.Format(" Hand calculation (A = {0:0.00} Ci): {1:0.00} cGy", currentActivity / 1000, sumDose) + System.Environment.NewLine;
            else message += String.Format(" Hand calculation (A = {0:0.00} Ci): {1:0.00} cGy", plannedActivity / 1000, sumDose) + System.Environment.NewLine;
            double percentDiff = 100 * ((sumDose - QAptDose) / QAptDose);
            message += String.Format(" Percent Difference: {0:0.0}%", percentDiff) + System.Environment.NewLine;
            message += String.Format(" Tolerance: {0}%", doseAgreeTolerance) + System.Environment.NewLine;
            message += String.Format(" Within tolerance?: {0}", Math.Abs(percentDiff) <= doseAgreeTolerance ? "YES" : "NO") + System.Environment.NewLine;
            doseCalcResults.Text = message;
            doseCalcScroller.ScrollToBottom();
            return false;
        }

        private void writeExcel_Click(object sender, RoutedEventArgs e)
        {
            RadioactiveSource source = plan.Catheters.First().TreatmentUnit.GetActiveRadioactiveSource();
            if (source == null) { MessageBox.Show("No radioactive source present in catheter!!"); return; }

            //get the calibration date of the source, today's date, and the interval between the two
            DateTime? calDate = source.CalibrationDate;
            DateTime today = DateTime.Now;
            TimeSpan interval;
            //get the interval between the source calibration date and today. Set to zero if calDate dose NOT have a value
            if (calDate.HasValue) interval = today.Subtract((DateTime)calDate);
            else interval = today.Subtract(today);

            //calculate current source activity. NOTE: THE ACITIVTY CALCULATION IS PER DAY AT 12:00 AM! This matches the decay calculation in Eclipse. interval.days is an integer. Must be cast to a double
            List<Catheter> applicators = plan.Catheters.ToList();

            // create Excel App (similar to WriteResultsExcel_Click method)
            Excel.Application myExcelApplication = new Excel.Application();
            Excel.Workbook myExcelWorkbook;
            Excel.Worksheet myExcelWorkSheet;
            myExcelApplication.DisplayAlerts = false; // turn off alerts

            //asign the proper data and start rows depending if needles were used or not
            int doseToQApointRow = -1;
            int sourceActivityRow = 8;
            int totalTimeRow = 11;
            List<int> startDataRow = new List<int> { };

            //if (applicators.Count == 1)
            //{
            //    //one channel --> likely vaginal cylinder. Grab the cylinder second check spreadsheet
            //    if (secondCheckFile == "") secondCheckFile = "Cylinder 2nd Check.xlsx";
            //    //start rows of where to place the data
            //    sourceActivityRow = 8;
            //    doseToQApointRow = 15;
            //    startDataRow.Add(26);
            //}
            if (applicators.Count >= 1)
            {
                //multiple channels --> interstitial needles were likely used. Need to use the interstitial second check spreadsheet. This will also be used if the treatment was a TnO
                if (secondCheckFile == "") secondCheckFile = "Interstitial 2nd Check.xlsx";
                doseToQApointRow = 18;
                //start rows of where to place the dwell position and time data for each applicator/needle
                for (int i = 0; i < applicators.Count; i++) startDataRow.Add(29 + i * 25);
            }
            //no applicators present in plan
            else return;

            if (!File.Exists(System.IO.Path.Combine(patientDataBase, secondCheckFile))) { MessageBox.Show(string.Format("Error! The specified Excel file:\n{0}\ndoes not exist! Exiting", System.IO.Path.Combine(patientDataBase, secondCheckFile))); return; }
            // open the existing excel file
            myExcelWorkbook = (Excel.Workbook)(myExcelApplication.Workbooks._Open(System.IO.Path.Combine(patientDataBase, secondCheckFile),
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value));

            //get the first worksheet
            myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1];

            //start adding data to the worksheet (patient name and today's date)
            myExcelWorkSheet.Cells[1, "B"] = VMS.TPS.Script.GetScriptContext().Patient.Name;
            myExcelWorkSheet.Cells[4, "B"] = DateTime.Now.ToString();
            if (applicators.Count >= 1)
            {
                //for whatever reason, they want today's date entered two more times on the interstitial spreadsheet
                myExcelWorkSheet.Cells[11, "D"] = DateTime.Now.ToString();
                myExcelWorkSheet.Cells[28, "C"] = DateTime.Now.ToString();
                //enter the current source activity
                if(useCurrentActivity) myExcelWorkSheet.Cells[sourceActivityRow, "A"] = String.Format("{0:0.000}", currentActivity / 1000);
                else myExcelWorkSheet.Cells[sourceActivityRow, "A"] = String.Format("{0:0.000}", plannedActivity / 1000);
            }
            //else myExcelWorkSheet.Cells[sourceActivityRow, "A"] = String.Format("{0:0.000}", sourceActivity);

            //write the dose to the QA point to the spreadsheet
            myExcelWorkSheet.Cells[doseToQApointRow + 3, "A"] = String.Format("{0:0.00}", QAptDose);

            myExcelWorkSheet.Cells[doseToQApointRow, "B"] = String.Format("{0:0.00}", QAptLoc.x / 10);
            myExcelWorkSheet.Cells[doseToQApointRow, "C"] = String.Format("{0:0.00}", QAptLoc.y / 10);
            myExcelWorkSheet.Cells[doseToQApointRow, "D"] = String.Format("{0:0.00}", QAptLoc.z / 10);


            //index denotes the current applicator number. count denotes the current dwell point for a given applicator
            int index = 0;
            int count = 0;
            //iterate through each catheter and each source position within a given catheter and write the relevant data
            foreach (Catheter c in applicators)
            {
                List<SourcePosition> sourcePos = c.SourcePositions.ToList();
                foreach (SourcePosition s in sourcePos)
                {
                    //the position of the center of the source (in user coordinates!)
                    VVector pos = s.Translation;
                    pos = plan.StructureSet.Image.DicomToUser(pos, plan);
                    //if (applicators.Count == 1)
                    //{
                    //    // message += String.Format("{0:0.0}, {1:0.00}, {2:0.00}, {3:0.00}\n", s.DwellTime, pos.x / 10, pos.y / 10, pos.z / 10);
                    //    myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "B"] = String.Format("{0:0.0}", s.DwellTime);
                    //    myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "C"] = String.Format("{0:0.00}", pos.x / 10);
                    //    myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "D"] = String.Format("{0:0.00}", pos.y / 10);
                    //    myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "E"] = String.Format("{0:0.00}", pos.z / 10);
                    //}
                    //else
                    //{
                        //the pullback distance from the tip of the applicator (specifically requested for the interstitial and TnO cases)
                        myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "B"] = String.Format("{0:0.0}", (c.ApplicatorLength - c.GetSourcePosCenterDistanceFromTip(s)) / 10);
                        myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "C"] = String.Format("{0:0.0}", s.DwellTime);
                        myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "D"] = String.Format("{0:0.00}", pos.x / 10);
                        myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "E"] = String.Format("{0:0.00}", pos.y / 10);
                        myExcelWorkSheet.Cells[startDataRow.ElementAt(index) + count, "F"] = String.Format("{0:0.00}", pos.z / 10);
                    //}
                    count++;
                }
                //increment the current catheter number
                index++;
                //reset the current dwell point to 0
                count = 0;
            }
            //write the total treatment time
            myExcelWorkSheet.Cells[totalTimeRow, "A"] = String.Format("{0:0.0}", totalPlannedTime);
            //in our current workflow, these two items will be the same (might not be if they were planned and treated on separate days. Use caution)
            if (applicators.Count >= 1) myExcelWorkSheet.Cells[totalTimeRow + 3, "A"] = String.Format("{0:0.0}", totalPlannedTime);

            string result = new helpers().WriteResultsToExcel(patientDataBase, secondCheckFile, myExcelWorkbook);
            if(result != "") doseCalcResults.Text += result;
        }

        private void writeText_Click(object sender, RoutedEventArgs e)
        {
            string fileName = new helpers().WriteResultsText(patientDataBase, doseCalcResults.Text);
            if (fileName != "")
            {
                doseCalcResults.Text += System.Environment.NewLine + String.Format("Results written to txt file: {0}", fileName.Substring(fileName.LastIndexOf("\\") + 1, fileName.Length - fileName.LastIndexOf("\\") - 1)) + System.Environment.NewLine;
                doseCalcScroller.ScrollToBottom();
            }
        }
    }
}
