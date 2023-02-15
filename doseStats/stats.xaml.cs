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
using System.IO;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using Excel = Microsoft.Office.Interop.Excel;
using Structure = VMS.TPS.Common.Model.API.Structure;
using StructureSet = VMS.TPS.Common.Model.API.StructureSet;
using Course = VMS.TPS.Common.Model.API.Course;
//using AriaQ_v15;

namespace doseStats
{
    //greek characters
    // α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω
    // Α Β Γ Δ Ε Ζ Η Θ Ι Κ Λ Μ Ν Ξ Ο Π Ρ Σ Τ Υ Φ Χ Ψ Ω
    public partial class stats : Window
    {
        //the brachy plan open in the current context
        BrachyPlanSetup plan = null;
        //the structure set associated with the brachy plan open in the current context
        StructureSet selectedSS = null;
        //Excel template where the second check data should be written
        string secondCheckFile = "";
        //list of treatment approved plans and the brachy plan open in the current context
        List<BrachyPlanSetup> plans = new List<BrachyPlanSetup> { };
        //a separate array for the requeted statistics (in case the user wants to query additional statistics beyond the defaults)
        List<Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>> statsRequest;
        //the obtained statistics from the plans
        //structure id, alpha/beta, statistic requested, query value, units, result
        List<Tuple<string, double, List<Tuple<string, double, string, List<double>>>>> statsResults = new List<Tuple<string, double, List<Tuple<string, double, string, List<double>>>>> { };
        //the data for each structure that will be written to the excel file. Unfortunately, I have to use this messy format since there is no structure to how the data is written to the excel file (except fraction # increases with column #)
        List<List<double>> excelData = new List<List<double>> { };
        //instance of the struct to copy the configuration parameters to the current class
        VMS.TPS.Script.Parameters p;
        //number of BRACHY fractions
        int numFractions;
        //external beam Rx dose
        double EBRTRxDose;
        //this flag is used to signal if an added requested statistic is the first in the list (used for proper addition/remove of the header)
        bool firstStatStruct = true;
        //this flag is used to signal if you want to assume or retrieve the dose from the external beam plan in the calculation of EQD2 (this option was removed from the GUI following Dr. Kidd's request)
        bool assumeMaxEQD2 = true;
        int clearStatBtnCounter = 0;
        //this flag is used to signal if the plan Id's are in the correct format. Specifically, if the fourth character in the plan Id is NOT an integer, this flag is set to true indicating there was a problem
        public bool formatError = false;
        //flag used to indicate if the ideal doses are shown
        bool isIdealDoses = false;
        //scaling factor to account for it the user entered > 1 fraction in the plan properties
        int scaleFactor;

        public stats(BrachyPlanSetup brachyPlan, VMS.TPS.Script.Parameters config)
        {
            InitializeComponent();
            p = config;
            //add new empty lists to the excel data lists (one for each requested statistic)
            for (int i = 0; i < p.excelStatistics.Count(); i++) excelData.Add(new List<double> { });
            plan = brachyPlan;
            scaleFactor = (int)plan.NumberOfFractions;
            selectedSS = plan.StructureSet;
            //determine the number of BRACHY fractions from the primary reference point in the open plan in the current context. This method is required since there is no other information in the plan that identifies the number of fractions
            try { numFractions = (int)(plan.PrimaryReferencePoint.TotalDoseLimit / plan.PrimaryReferencePoint.SessionDoseLimit); }
            catch { MessageBox.Show(String.Format("Primary reference point dose limits not set! \nPlease fix and try again!")); formatError = true; return; }
            //additional logic for interstitial treatments that are typically planned 2 fractions at a time (these treatments are typically 4 fractions)
            if (plan.ProtocolID.ToLower().Contains("interstitial") && numFractions == 2) numFractions = 4;

            //get treatment approved plans. Add plan open in current context to stack if it is NOT treatment approved. Check that the plan Id's in the stack have the correct formatting
            plans = plan.Course.BrachyPlanSetups.Where(x => x.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved).ToList();
            if (plan.ApprovalStatus != PlanSetupApprovalStatus.TreatmentApproved) plans.Add(plan);
            foreach (BrachyPlanSetup bp in plans)
            {
                int loc = bp.Id.IndexOf("HDR");
                if (loc != -1)
                {
                    if (!int.TryParse(bp.Id.Substring(loc + 3, 1), out int num))
                    {
                        MessageBox.Show(String.Format("Error! The character following 'HDR' is NaN! \nPlan: {0}\n\nPlease fix and try again!", bp.Id));
                        formatError = true;
                        return;
                    }
                }
                else if (bp.Id.Length - bp.Id.Replace("_", "").Length != 2)
                {
                    MessageBox.Show(String.Format("Error! The string 'HDR' is NOT present in the plan Id and there are not two underscore characters! \nPlan: {0}\n\nOnly second dose calculation check possible!", bp.Id));
                    formatError = true;
                    return;
                }
            }

            bool containsHDR = plans.All(x => x.Id.Contains("HDR"));
            //sort the plans (all plans should start with 'HDR' and the next character should be the fraction number, which will be used to sort the plans)
            if (containsHDR) plans.Sort(delegate (BrachyPlanSetup x, BrachyPlanSetup y) { return x.Id.Substring(x.Id.IndexOf("HDR") + 3, 1).CompareTo(y.Id.Substring(y.Id.IndexOf("HDR") + 3, 1)); });
            else plans.Sort(delegate (BrachyPlanSetup x, BrachyPlanSetup y)
            { return double.Parse(x.Id.Substring(x.Id.IndexOf("_") + 1, x.Id.IndexOf("_", x.Id.IndexOf("_") + 1) - x.Id.IndexOf("_") - 1)).CompareTo(double.Parse(y.Id.Substring(y.Id.IndexOf("_") + 1, y.Id.IndexOf("_", y.Id.IndexOf("_") + 1) - y.Id.IndexOf("_") - 1))); });

            //if (plan.ProtocolID.Contains("T&O") || plan.ProtocolID.Contains("TO")) 
            //    statsRequest.Add(new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>("pt A", 10.0, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>> { new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Dmean (Gy)", 0.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute) }));

            //update the GUI with the number of EBRT fractions and dose per fraction
            EBRTdosePerFxTB.Text = p.EBRTdosePerFx.ToString();
            EBRTnumFxTB.Text = p.EBRTnumFx.ToString();
            //assign defaultStats array to statsRequest, retrieve these statistics from the plans, update the GUI, then add the default stats request to the list of requested statistics in the GUI (least important, which is why it is last)
            statsRequest = new List<Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>>(p.defaultStats);
            getStatsFromPlans();
            updateStats();
            add_stat_volumes(statsRequest);

            //create and bind keyboard macros to make script execution faster
            runStaticsMacro.InputGestures.Add(new KeyGesture(Key.E, ModifierKeys.Control));
            writeExcelMacro.InputGestures.Add(new KeyGesture(Key.W, ModifierKeys.Control));
            runSecondDoseCalcMacro.InputGestures.Add(new KeyGesture(Key.D, ModifierKeys.Control));
            closeScriptMacro.InputGestures.Add(new KeyGesture(Key.Q, ModifierKeys.Control));
           // toggleAssumeMaxEQD2.InputGestures.Add(new KeyGesture(Key.A, ModifierKeys.Control));
            showHelpMacro.InputGestures.Add(new KeyGesture(Key.H, ModifierKeys.Control));
            toggleShowIdealDoses.InputGestures.Add(new KeyGesture(Key.A, ModifierKeys.Control));
            openManualAdjustWindow.InputGestures.Add(new KeyGesture(Key.M, ModifierKeys.Control));

            CommandBindings.Add(new CommandBinding(runStaticsMacro, calculateStatistics_Click));
            CommandBindings.Add(new CommandBinding(writeExcelMacro, WriteResultsExcel_Click));
            CommandBindings.Add(new CommandBinding(runSecondDoseCalcMacro, runSecondCheck_Click));
            CommandBindings.Add(new CommandBinding(closeScriptMacro, closeWindow));
            CommandBindings.Add(new CommandBinding(showHelpMacro, openHelp_Click));
            //CommandBindings.Add(new CommandBinding(toggleAssumeMaxEQD2, toggleEQD2));
            CommandBindings.Add(new CommandBinding(toggleShowIdealDoses, toggleIdealDosesCheckBox));
            CommandBindings.Add(new CommandBinding(openManualAdjustWindow, showMDAwindow));
        }

        RoutedCommand runStaticsMacro = new RoutedCommand();
        RoutedCommand writeExcelMacro = new RoutedCommand();
        RoutedCommand runSecondDoseCalcMacro = new RoutedCommand();
        RoutedCommand closeScriptMacro = new RoutedCommand();
        //RoutedCommand toggleAssumeMaxEQD2 = new RoutedCommand();
        RoutedCommand toggleShowIdealDoses = new RoutedCommand();
        RoutedCommand openManualAdjustWindow = new RoutedCommand();
        RoutedCommand showHelpMacro = new RoutedCommand();

        //removed per Dr. Kidd's request
        private void toggleEQD2(object sender, EventArgs e)
        {
            if (assumeMaxEQD2_ckbox.IsChecked.Value) assumeMaxEQD2_ckbox.IsChecked = false;
            else assumeMaxEQD2_ckbox.IsChecked = true;
            updateAssumeEQD2Option();
        }

        //removed per Dr. Kidd's request
        private void updateAssumeEQD2Option()
        {
            if (assumeMaxEQD2_ckbox.IsChecked.Value)
            {
                assumeMaxEQD2 = true;
                EBRTdosePerFxTB.Text = String.Format("{0}", 1.8);
                EBRTnumFxTB.Text = String.Format("{0}", 27);
            }
            else assumeMaxEQD2 = false;

            //clear the existing vectors
            statsResults.Clear();
            //clear excel data vectors
            for (int i = 0; i < excelData.Count; i++) excelData.ElementAt(i).Clear();
            getStatsFromPlans();
            updateStats();
        }

        //removed per Dr. Kidd's request
        private void assumeMaxEQD2_ckbox_Click(object sender, RoutedEventArgs e)
        {
            //updateAssumeEQD2Option(); 
        }

        //close the open dose statistics window. Called from the Ctrl + Q keyboard macro
        private void closeWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void openShortcuts_Click(object sender, RoutedEventArgs e)
        {
            string message = "Some useful keyboard shortcuts:" + System.Environment.NewLine;
            message += "Quit script --------------> Ctrl + Q" + System.Environment.NewLine;
            message += "Calculate statistics --------------> Ctrl + E" + System.Environment.NewLine;
            message += "Write results to Excel --------------> Ctrl + W" + System.Environment.NewLine;
            message += "Run second dose calc --------------> Ctrl + D" + System.Environment.NewLine;
            message += "Display help guide --------------> Ctrl + H" + System.Environment.NewLine;
            message += "Toggle show ideal doses --------------> Ctrl + A" + System.Environment.NewLine;
            message += "Open manual dose adjustment window --------------> Ctrl + M" + System.Environment.NewLine;
            MessageBox.Show(message);
        }

        //open the documentation for this script (PDF file)
        private void openHelp_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(p.documentation);
        }

        private void showMDAwindow(object sender, RoutedEventArgs e)
        {
            manualDoseAdjustment mdaWindow = new manualDoseAdjustment(numFractions, statsResults, p);
            mdaWindow.ShowDialog();
        }

        private void toggleIdealDosesCheckBox(object sender, RoutedEventArgs e)
        {
            if (showIdealsCheckBox.IsChecked.Value) showIdealsCheckBox.IsChecked = false;
            else showIdealsCheckBox.IsChecked = true;
        }

        private void showIdealDoses(object sender, RoutedEventArgs e)
        {
            if (showIdealsCheckBox.IsChecked.Value)
            {
                isIdealDoses = true;
                //no external beam plan found or we want to assume the max EQD2. Assume alpha/beta ratios of 10 Gy and 3 Gy for the tumor and OAR, respectively
                //variables to hold the EBRT total for the target and the OARs
                double tumorEBRTtotal = p.EBRTnumFx * p.EBRTdosePerFx * (p.EBRTdosePerFx + 10) / (2.0 + 10.0);
                double bladderEBRTtotal = 0.0, bowelEBRTtotal = 0.0, rectumEBRTtotal = 0.0, sigmoidEBRTtotal = 0.0;
                bladderEBRTtotal = bowelEBRTtotal = rectumEBRTtotal = sigmoidEBRTtotal = p.EBRTnumFx * p.EBRTdosePerFx * (p.EBRTdosePerFx + 3) / (2.0 + 3.0);
                List<Tuple<string, double, List<Tuple<string, double, string, List<double>>>>> statsResults_temp = new List<Tuple<string, double, List<Tuple<string, double, string, List<double>>>>> { };
                List<Tuple<string, double, string, List<double>>> stats_temp = new List<Tuple<string, double, string, List<double>>> { };
                helpers h = new helpers();

                foreach (Tuple<string, double, List<Tuple<string, double, string, List<double>>>> itr in statsResults)
                {
                    //statistics, query value, units, vector of results from the plans
                    foreach (Tuple<string, double, string, List<double>> itr1 in itr.Item3)
                    {
                        if (itr1.Item1 == "Dose at Volume (Gy)" || itr1.Item1 == "Dmean (Gy)")
                        {
                            //need to compute EQD2 values for the requested absolute doses
                            double HDRsum = 0.0;
                            double HDR_EBRT_sum = 0.0;
                            double val = 0.0;
                            double currentHDREQD2 = 0.0;
                            for (int i = 0; i < numFractions; i++)
                            {
                                //EQD2 is calculated for each structure for a SINGLE fraction (i.e., the current fraction). The single fraction EQD2 values are then added to obtain the cumulative EQD2. 
                                //If the current loop iteration is less than the size of the results vector, calculate EQD2 based on the value in the results vector, otherwise, propagate the last element in the vector
                                //forward to the remaining fractions
                                if (i < itr1.Item4.Count) { val = itr1.Item4.ElementAt(i) * ((itr1.Item4.ElementAt(i) + itr.Item2) / (2 + itr.Item2)); if(i < itr1.Item4.Count - 1) currentHDREQD2 += val; }
                                else val = itr1.Item4.Last() * ((itr1.Item4.Last() + itr.Item2) / (2 + itr.Item2));
                                HDRsum += val;
                            }
                            //calculate the total EQD2 including HDR and EBRT
                            if (itr.Item1.Contains("gtv") || itr.Item1.Contains("ctv") || itr.Item1.Contains("pt A")) HDR_EBRT_sum = HDRsum + tumorEBRTtotal;
                            else
                            {
                                double total = 0.0;
                                //messy, but legacy code leftover from retrieving the EQD2 data from the actual external beam plans rather than assuming a particular dose was delivered
                                if (itr.Item1.Contains("bladder")) total = bladderEBRTtotal;
                                else if (itr.Item1.Contains("bowel")) total = bowelEBRTtotal;
                                else if (itr.Item1.Contains("rectum")) total = rectumEBRTtotal;
                                else if (itr.Item1.Contains("sigmoid")) total = sigmoidEBRTtotal;
                                HDR_EBRT_sum = HDRsum + total;
                            }

                            //this is the logic to determine if the requested statistic has a dosimetric aim and/or limit that we are shooting for. Currently we have aims and limits for:
                            //Bladder D2cc, Bowel D2cc, Rectum D2cc, CTV D98%, CTV D90%, and PtA Dmean
                            //See the Gyn HDR BT spreadsheet for a list of current aims and limits
                            bool met = false;
                            Tuple<string, string> value = h.getAimLimit(p, itr.Item1, itr1.Item1, itr1.Item2, itr1.Item3);
                            string aim = value.Item1;
                            string limit = value.Item2;

                            //if either the aim or limit are nonempty, add these values to the reporting window text, close the final bracket, add the text to the window, and add a new line
                            if (aim != "" || limit != "")
                            {
                                met = h.checkIsMet(aim, limit, HDR_EBRT_sum);
                                if(!met)
                                {
                                    int deliveredFx = itr1.Item4.Count;
                                    itr1.Item4.RemoveAt(itr1.Item4.Count - 1);
                                    string aimLimitTemp;
                                    if (aim != "") aimLimitTemp = aim;
                                    else aimLimitTemp = limit;
                                    if (double.TryParse(aimLimitTemp.Substring(1, 2), out double eqd2Val))
                                    {
                                        if (itr.Item1.Contains("gtv") || itr.Item1.Contains("ctv") || itr.Item1.Contains("pt A")) eqd2Val -= (tumorEBRTtotal + currentHDREQD2);
                                        else eqd2Val -= (bladderEBRTtotal + currentHDREQD2);
                                        if (aimLimitTemp.Substring(0, 1) == ">") eqd2Val += 0.01;
                                        else eqd2Val -= 0.01;
                                        itr1.Item4.Add((Math.Sqrt(Math.Pow(itr.Item2, 2) + 4 * eqd2Val * (2 + itr.Item2) / (numFractions - deliveredFx + 1)) - itr.Item2) / 2);
                                        //MessageBox.Show(String.Format("{0}, {1:0.00}, {2:0.00}, {3}", itr.Item1, currentHDREQD2, eqd2Val, (numFractions - deliveredFx + 1)));
                                    }
                                    else MessageBox.Show(String.Format("Double value could not be parsed from Aims and Limits! ({0})", aimLimitTemp)); ;
                                }
                            }
                        }
                    }
                }
                updateStats();
                FlowDocument myFlowDoc = new FlowDocument();
                results.Text += (System.Environment.NewLine);
                results.Text += ("WARNING!! THE ABOVE DOSES REPRESENT THE IDEAL DOSES THAT JUST MEET THE EQD2 AIMS/LIMITS!" + System.Environment.NewLine);
                results.Text += ("THEY DO NOT REPRESENT THE ACHIEVED DOSES IN THE CURRENT PLAN!!!");
                resultsScroller.ScrollToBottom();
            }
            else
            {
                isIdealDoses = false;
                //clear the existing vectors
                statsResults.Clear();
                //clear excel data vectors
                for (int i = 0; i < excelData.Count; i++) excelData.ElementAt(i).Clear();
                getStatsFromPlans();
                updateStats();
            }
        }

        //code used to update the total EBRT dose dynamically as the text is changed
        private void EBRTdosePerFxTB_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!double.TryParse(EBRTdosePerFxTB.Text, out double dummy)) EBRTRxDoseTB.Text = "";
            else if (dummy <= 0.0)
            {
                MessageBox.Show("Error! The EBRT dose per fraction must be a number and non-negative!");
                EBRTRxDoseTB.Text = "";
            }
            else resetEBRTRxDose();
        }

        //code used to update the total EBRT dose dynamically as the text is changed
        private void EBRTnumFxTB_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!int.TryParse(EBRTnumFxTB.Text, out int dummy)) EBRTRxDoseTB.Text = "";
            else if (dummy < 1)
            {
                MessageBox.Show("Error! The EBRT number of fractions must be a integer and >= 1!");
                EBRTRxDoseTB.Text = "";
            }
            else resetEBRTRxDose();
        }

        //update the total EBRT dose if the EBRT dose per fraction and EBRT number of fractions text values are in the correct format. In addition, update the EBRT dose per fraction and number of fractions in the code
        private void resetEBRTRxDose()
        {
            if (double.TryParse(EBRTdosePerFxTB.Text, out double temp1) && int.TryParse(EBRTnumFxTB.Text, out int temp2))
            {
                p.EBRTdosePerFx = temp1; p.EBRTnumFx = temp2;
                EBRTRxDose = (p.EBRTdosePerFx * p.EBRTnumFx);
                EBRTRxDoseTB.Text = EBRTRxDose.ToString();
            }
        }

        //clear the entire list of requested statistics
        private void clear_stats_Click(object sender, RoutedEventArgs e) { clear_statistic_parameters_list(); }

        //clear the entire list of requested statistics
        private void clear_statistic_parameters_list()
        {
            firstStatStruct = true;
            stat_parameters.Children.Clear();
            clearStatBtnCounter = 0;
        }

        //the code used to add the header to the list of requested statistics
        private void add_stat_header()
        {
            //structure id, alpha/beta, statistic requested, query value, Volume representation (relative or absolute)
            StackPanel sp1 = new StackPanel();
            sp1.Height = 30;
            sp1.Width = stat_parameters.Width;
            sp1.Orientation = Orientation.Horizontal;
            sp1.Margin = new Thickness(5, 0, 5, 5);

            Label strName = new Label();
            strName.Content = "Structure";
            strName.HorizontalAlignment = HorizontalAlignment.Center;
            strName.VerticalAlignment = VerticalAlignment.Top;
            strName.Width = 100;
            strName.FontSize = 14;
            strName.Margin = new Thickness(30, 0, 0, 0);

            Label spareType = new Label();
            spareType.Content = "α/β (Gy)";
            spareType.HorizontalAlignment = HorizontalAlignment.Center;
            spareType.VerticalAlignment = VerticalAlignment.Top;
            spareType.Width = 60;
            spareType.FontSize = 14;
            spareType.Margin = new Thickness(0, 0, 0, 0);

            Label volLabel = new Label();
            volLabel.Content = "Statistic";
            volLabel.HorizontalAlignment = HorizontalAlignment.Center;
            volLabel.VerticalAlignment = VerticalAlignment.Top;
            volLabel.Width = 110;
            volLabel.FontSize = 14;
            volLabel.Margin = new Thickness(29, 0, 0, 0);

            Label doseLabel = new Label();
            doseLabel.Content = "Value";
            doseLabel.HorizontalAlignment = HorizontalAlignment.Center;
            doseLabel.VerticalAlignment = VerticalAlignment.Top;
            doseLabel.Width = 60;
            doseLabel.FontSize = 14;
            doseLabel.Margin = new Thickness(3, 0, 0, 0);

            Label priorityLabel = new Label();
            priorityLabel.Content = "Units";
            priorityLabel.HorizontalAlignment = HorizontalAlignment.Center;
            priorityLabel.VerticalAlignment = VerticalAlignment.Top;
            priorityLabel.Width = 65;
            priorityLabel.FontSize = 14;
            priorityLabel.Margin = new Thickness(15, 0, 0, 0);

            sp1.Children.Add(strName);
            sp1.Children.Add(spareType);
            sp1.Children.Add(volLabel);
            sp1.Children.Add(doseLabel);
            sp1.Children.Add(priorityLabel);
            stat_parameters.Children.Add(sp1);

            firstStatStruct = false;
        }

        //add a requested statistic to the list
        private void add_stat_Click(object sender, RoutedEventArgs e)
        {
            if (firstStatStruct) add_stat_header();
            add_stat_volumes(new List<Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>> { new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>("--select--", 0.0, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>> { new Tuple<string, double, VolumePresentation, DoseValuePresentation>("--select--", 0.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute) }) });
            statParamScroller.ScrollToBottom();
        }

        //add default list of requested statistics
        private void add_defaults_Click(object sender, RoutedEventArgs e)
        {
            //if (plan.ProtocolID.Contains("T&O") || plan.ProtocolID.Contains("TO"))
            //statsRequest.Add(new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>("pt A", 10.0, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>> { new Tuple<string, double, VolumePresentation, DoseValuePresentation>("Dmean (Gy)", 0.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute) }));
            clear_statistic_parameters_list();
            statsRequest = new List<Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>>(p.defaultStats);
            add_stat_volumes(statsRequest);
            statParamScroller.ScrollToBottom();
        }

        //generic code used to add requested statistics to the list. First, a check is performed to see if this is the first requested statistics added to the list. If so, the header is added first. This code is used whether a blank requested statistic is added or if 
        //the default list of requested statistics is added
        private void add_stat_volumes(List<Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>> defaultList)
        {
            // structure id, alpha / beta, statistic requested, query value, Volume representation(relative or absolute), dose representation (relative or absolute)
            if (firstStatStruct) add_stat_header();
            foreach (Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>> itr in defaultList)
            {
                foreach (Tuple<string, double, VolumePresentation, DoseValuePresentation> itr1 in itr.Item3)
                {
                    StackPanel sp = new StackPanel();
                    sp.Height = 30;
                    sp.Width = stat_parameters.Width;
                    sp.Orientation = Orientation.Horizontal;
                    sp.Margin = new Thickness(5);

                    ComboBox stat_str_cb = new ComboBox();
                    stat_str_cb.Name = "stat_str_cb";
                    stat_str_cb.Width = 120;
                    stat_str_cb.Height = sp.Height - 5;
                    stat_str_cb.HorizontalAlignment = HorizontalAlignment.Left;
                    stat_str_cb.VerticalAlignment = VerticalAlignment.Top;
                    stat_str_cb.Margin = new Thickness(5, 5, 0, 0);

                    stat_str_cb.Items.Add("--select--");
                    foreach (string s in p.structures) stat_str_cb.Items.Add(s);
                    stat_str_cb.Text = itr.Item1;
                    stat_str_cb.HorizontalContentAlignment = HorizontalAlignment.Center;
                    stat_str_cb.SelectionChanged += new SelectionChangedEventHandler(stat_str_cb_change);
                    sp.Children.Add(stat_str_cb);

                    TextBox ab_tb = new TextBox();
                    ab_tb.Name = "ab_tb";
                    ab_tb.Width = 50;
                    ab_tb.Height = sp.Height - 5;
                    ab_tb.HorizontalAlignment = HorizontalAlignment.Left;
                    ab_tb.VerticalAlignment = VerticalAlignment.Center;
                    ab_tb.Margin = new Thickness(5, 5, 0, 0);
                    if (itr.Item1 == "ctv" || itr.Item1 == "pt A" || itr.Item1 == "gtv") ab_tb.Text = String.Format("{0:0.0}", 10.0);
                    else if (itr.Item1 != "--select--") ab_tb.Text = String.Format("{0:0.0}", 3.0);
                    else ab_tb.Text = String.Format("{0:0.0}", 0.0);
                    //ab_tb.Text = String.Format("{0:0.0}", itr.Item2);
                    ab_tb.TextAlignment = TextAlignment.Center;
                    ab_tb.IsReadOnly = true;
                    sp.Children.Add(ab_tb);

                    ComboBox statType_cb = new ComboBox();
                    statType_cb.Name = "type_cb";
                    statType_cb.Width = 140;
                    statType_cb.Height = sp.Height - 5;
                    statType_cb.HorizontalAlignment = HorizontalAlignment.Left;
                    statType_cb.VerticalAlignment = VerticalAlignment.Top;
                    statType_cb.Margin = new Thickness(5, 5, 0, 0);
                    string[] types = new string[] { "--select--", "Dose at Volume (Gy)", "Dose at Volume (%)", "Dmean (Gy)", "Dmean (%)", "Volume at Dose (cc)", "Volume at Dose (%)", "Volume (cc)" };
                    foreach (string s in types) statType_cb.Items.Add(s);
                    statType_cb.Text = itr1.Item1;
                    statType_cb.HorizontalContentAlignment = HorizontalAlignment.Center;
                    statType_cb.SelectionChanged += new SelectionChangedEventHandler(statType_cb_change);
                    sp.Children.Add(statType_cb);

                    TextBox value_tb = new TextBox();
                    value_tb.Name = "value_tb";
                    value_tb.Width = 50;
                    value_tb.Height = sp.Height - 5;
                    value_tb.HorizontalAlignment = HorizontalAlignment.Left;
                    value_tb.VerticalAlignment = VerticalAlignment.Center;
                    value_tb.Margin = new Thickness(5, 5, 0, 0);
                    value_tb.Text = String.Format("{0:0.0}", itr1.Item2);
                    value_tb.TextAlignment = TextAlignment.Center;
                    if (statType_cb.SelectedItem.ToString().Contains("Dmean") || statType_cb.SelectedItem.ToString().Contains("Volume (cc)")) value_tb.Visibility = Visibility.Hidden;
                    sp.Children.Add(value_tb);

                    ComboBox unit_cb = new ComboBox();
                    unit_cb.Name = "unit_cb";
                    unit_cb.Width = 100;
                    unit_cb.Height = sp.Height - 5;
                    unit_cb.HorizontalAlignment = HorizontalAlignment.Left;
                    unit_cb.VerticalAlignment = VerticalAlignment.Top;
                    unit_cb.Margin = new Thickness(5, 5, 0, 0);
                    unit_cb.HorizontalContentAlignment = HorizontalAlignment.Center;

                    if (statType_cb.SelectedItem.ToString().Contains("Dmean") || statType_cb.SelectedItem.ToString().Contains("Volume (cc)")) unit_cb.Visibility = Visibility.Hidden;
                    else
                    {
                        List<string> unitType = new List<string> { "--select--", "%" };
                        if (itr1.Item1.Contains("Dose at Volume")) unitType.Add("cc");
                        else unitType.Add("Gy");
                        foreach (string s in unitType) unit_cb.Items.Add(s);
                        unit_cb.Text = "%";
                        if (itr1.Item1 == "--select--") unit_cb.Text = "--select--";
                        if (itr1.Item1.Contains("Dose at Volume")) { if (itr1.Item3 == VolumePresentation.AbsoluteCm3) unit_cb.Text = "cc"; }
                        else if (itr1.Item1.Contains("Volume at Dose")) { if (itr1.Item4 == DoseValuePresentation.Absolute) unit_cb.Text = "Gy"; }
                    }
                    sp.Children.Add(unit_cb);

                    Button clearStatStructBtn = new Button();
                    clearStatBtnCounter++;
                    clearStatStructBtn.Name = "clearOptStructBtn" + clearStatBtnCounter;
                    clearStatStructBtn.Content = "Clear";
                    clearStatStructBtn.Click += new RoutedEventHandler(this.clearStatStructBtn_click);
                    clearStatStructBtn.Width = 50;
                    clearStatStructBtn.Height = sp.Height - 5;
                    clearStatStructBtn.HorizontalAlignment = HorizontalAlignment.Left;
                    clearStatStructBtn.VerticalAlignment = VerticalAlignment.Top;
                    clearStatStructBtn.Margin = new Thickness(10, 5, 0, 0);
                    sp.Children.Add(clearStatStructBtn);

                    stat_parameters.Children.Add(sp);
                }
            }
        }

        //code to dynamically adjust the alpha/beta ratio as the selected structure changes:
        //CTV and pt a --> alpha/beta = 10
        //OARs --> alpha/beta = 3
        //no structure selected --> 0
        private void stat_str_cb_change(object sender, System.EventArgs e)
        {
            ComboBox c = (ComboBox)sender;
            bool row = false;
            foreach (object obj in stat_parameters.Children)
            {
                foreach (object obj1 in ((StackPanel)obj).Children)
                {
                    if (row)
                    {
                        TextBox tb = (obj1 as TextBox);
                        if (c.SelectedItem.ToString().Contains("ctv") || c.SelectedItem.ToString().Contains("pt A") || c.SelectedItem.ToString().Contains("gtv")) tb.Text = String.Format("{0:0.0}", 10.0);
                        else if (!c.SelectedItem.ToString().Contains("--select--")) tb.Text = String.Format("{0:0.0}", 3.0);
                        else tb.Text = String.Format("{0:0.0}", 0.0);
                        return;
                    }
                    //the combobox has a unique tag to it, so we can just loop through all children in the stat_parameters children list and find which combobox is equivalent to our combobox. Once we find the right
                    //combobox, set the row flag to true. That way we can perform operations on the next child element (should be the alpha/beta textbox)
                    if (obj1.Equals(c)) row = true;
                }
            }
        }

        //code used to dynamically adjust the units for the requested statistic for a given structure as the requested statistic changes
        private void statType_cb_change(object sender, System.EventArgs e)
        {
            //not the most elegent code, but it works
            ComboBox c = (ComboBox)sender;
            bool row = false;
            bool delayedRow = false;
            foreach (object obj in stat_parameters.Children)
            {
                foreach (object obj1 in ((StackPanel)obj).Children)
                {
                    if (delayedRow)
                    {
                        //adjust the available items in the units combobox
                        ComboBox unit = (obj1 as ComboBox);
                        unit.Items.Clear();
                        //if Dmean or Volume are selected for the requested statistic, hide this combobox
                        if (c.SelectedItem.ToString().Contains("Dmean") || c.SelectedItem.ToString().Contains("Volume (cc)")) unit.Visibility = Visibility.Hidden;
                        else
                        {
                            unit.Visibility = Visibility.Visible;
                            if (c.SelectedItem.ToString() != "--select--")
                            {
                                //'--select--' and '%' will always be present as options in the list
                                unit.Items.Add("--select--");
                                unit.Items.Add("%");
                                //If a volume is requested, add 'cc' as an option. Otherwise, add 'Gy" as an option. Set the default to '--select--'
                                if (c.SelectedItem.ToString().Contains("Dose at Volume")) unit.Items.Add("cc");
                                else if (c.SelectedItem.ToString().Contains("Volume at Dose")) unit.Items.Add("Gy");
                                unit.Text = "--select--";
                            }
                        }
                        return;
                    }

                    if (row)
                    {
                        //set the delayed flag to true to indicate we want to perform operations on the next child element as well (should be the units combobox)
                        delayedRow = true;
                        //if Dmean or Volume are selected for the requested statistic, remove the Value text box (i.e., hide it). Otherwise, make it visible
                        if (c.SelectedItem.ToString().Contains("Dmean") || c.SelectedItem.ToString().Contains("Volume (cc)")) (obj1 as TextBox).Visibility = Visibility.Hidden;
                        else (obj1 as TextBox).Visibility = Visibility.Visible;
                    }
                    //the combobox has a unique tag to it, so we can just loop through all children in the stat_parameters children list and find which combobox is equivalent to our combobox. Set the row
                    //flag to true to perform operations on the next child element
                    if (obj1.Equals(c)) row = true;
                }
            }
        }

        private void clearStatStructBtn_click(object sender, EventArgs e)
        {
            //find the exact clear button that was hit
            Button btn = (Button)sender;
            int i = 0;
            int k = 0;
            foreach (object obj in stat_parameters.Children)
            {
                foreach (object obj1 in ((StackPanel)obj).Children)
                {
                    if (obj1.Equals(btn)) k = i;
                }
                if (k > 0) break;
                i++;
            }

            //clear entire list if there are only two entries (header + 1 real entry). Otherwise, remove the entire row of children at row k
            if (stat_parameters.Children.Count < 3) clear_statistic_parameters_list();
            else stat_parameters.Children.RemoveAt(k);
        }

        //the organization of this method is a bit messy, but it's necessary because we assume the list of requested statistics is unordered. For this script to operate properly, the list of requested statistics
        //must be ordered by the structure Id's in the requested statistics list (because of the way the arrays are constructed)
        private void calculateStatistics_Click(object sender, RoutedEventArgs e)
        {
            //if nothing is in the list of requested statistics, throw and error and return
            if (stat_parameters.Children.Count == 0)
            {
                MessageBox.Show("No dose statistics requested! Add some dose statistics and try again!");
                return;
            }

            //this list holds all the lines of data in requested parameter list. NOTE: THIS IS NOT THE FINAL REQUESTED STATISTIC LIST! The data will be sorted and the static information for each requested statistic
            //(i.e., structure id and alpha/beta ratio) will be consolidated so there is only one copy of the static information
            //structure id, alpha/beta, statistic requested, query value, Volume representation (relative or absolute)
            List<Tuple<string, double, string, double, string>> sortedList = new List<Tuple<string, double, string, double, string>> { };
            //variables to hold the requested info before adding each set of requests to the vector
            string structure = "";
            double ab = -1.0;
            string statType = "";
            double value = -1.0;
            string units = "";
            int combobxNum = 1;
            int txtbxNum = 1;
            bool headerObj = true;
            //cycle through all children
            foreach (object obj in stat_parameters.Children)
            {
                //skip over header row
                if (!headerObj)
                {
                    //loop over all children and parse the relevant structure, alpha/beta, requested statistic, statistic value, statistic units
                    foreach (object obj1 in ((StackPanel)obj).Children)
                    {
                        if (obj1.GetType() == typeof(ComboBox))
                        {
                            //if the combobox is visible, that means it has useful information that needs to be parsed
                            if ((obj1 as ComboBox).IsVisible)
                            {
                                //verify that something has been selected
                                string dummy = (obj1 as ComboBox).SelectedItem.ToString();
                                if (dummy == "--select--")
                                {
                                    MessageBox.Show("Error! \nStructure, statistic, or units not selected! \nSelect an option and try again");
                                    return;
                                }
                                //first combobox is the structure
                                if (combobxNum == 1) structure = dummy;
                                //second combobox is the requested statistic
                                else if (combobxNum == 2) statType = dummy;
                                //third combobox is the units on the requested statistic
                                else if (combobxNum == 3) units = dummy;
                            }
                            combobxNum++;
                        }
                        else if (obj1.GetType() == typeof(TextBox))
                        {
                            if (!string.IsNullOrWhiteSpace((obj1 as TextBox).Text))
                            {
                                //first text box is the alpha beta ratio (in Gy)
                                if (txtbxNum == 1) double.TryParse((obj1 as TextBox).Text, out ab);
                                //second text box is the statistic requested value
                                else if (txtbxNum == 2) double.TryParse((obj1 as TextBox).Text, out value);
                            }
                            txtbxNum++;
                        }
                    }
                    //if the selected structure is 'pt A', the user can only request the Dmean statistic
                    if (structure.Contains("pt A") && !statType.Contains("Dmean"))
                    {
                        MessageBox.Show("Error! \nOnly the Dmean statistic is compatable with the pt A structure! \nEnter new values and try again");
                        return;
                    }
                    //do some checks to ensure the integrity of the ab data and requested statistic value
                    if (ab == -1.0 || value == -1.0)
                    {
                        MessageBox.Show("Error! \nα/β or statistic value are invalid! \nEnter new values and try again");
                        return;
                    }
                    //if the row of data passes the above checks, add it the requested statistic list
                    else sortedList.Add(Tuple.Create(structure, ab, statType, value, units));
                    //reset the values of the variables used to parse the data
                    combobxNum = 1;
                    txtbxNum = 1;
                    ab = -1.0;
                    value = -1.0;
                }
                else headerObj = false;
            }

            //clear the existing vectors
            statsResults.Clear();
            statsRequest.Clear();
            for (int i = 0; i < excelData.Count; i++) excelData.ElementAt(i).Clear();

            //sort the list according to the structure Id's (i.e., item 1 in the sorted arrays)
            //List<Tuple<string, double, string, double, string>> sortedList = initialList.OrderBy(x => x.Item1).ToList();
            sortedList.Sort((x, y) => x.Item1.CompareTo(y.Item1));

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
                    statsRequest.Add(new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>(temp, temp2, new List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>(listTemp)));
                    //clear the vector of the dynamic attributes of the requested statistics. Update temp and temp2 to the current structure in the list
                    listTemp.Clear();
                    temp = itr.Item1;
                    temp2 = itr.Item2;
                }

                //grabbing the requested statistic value is easy enough, determining the requested units on that statistic requires a bit more logic. Also need to determine the units on the query statistic
                double stat = itr.Item4;
                DoseValuePresentation dvp;
                VolumePresentation vp;

                if (itr.Item3.Contains("Dose at Volume") || itr.Item3.Contains("Dmean"))
                {
                    //the query statistic is a volume or a mean dose
                    //determine if the query dose is absolute or relative
                    if (itr.Item3.Contains("Gy")) dvp = DoseValuePresentation.Absolute;
                    else dvp = DoseValuePresentation.Relative;
                    //determine the units of the requested volume if applicable
                    if (itr.Item5.Contains("cc")) vp = VolumePresentation.AbsoluteCm3;
                    else vp = VolumePresentation.Relative;
                }
                else
                {
                    //the query statistic is a dose or the volume of the structure was requested
                    //determine if the query volume is absolute or relative
                    if (itr.Item3.Contains("cc")) vp = VolumePresentation.AbsoluteCm3;
                    else vp = VolumePresentation.Relative;
                    //determine the units of the requested dose if applicable
                    if (itr.Item5.Contains("Gy")) dvp = DoseValuePresentation.Absolute;
                    else dvp = DoseValuePresentation.Relative;
                }
                //MessageBox.Show(string.Format("{0}, {1}, {2}, {3}, {4}, {5}", itr.Item1, itr.Item2, itr.Item3, stat, vp.ToString(), dvp.ToString()));
                //add a new entry to the requested statistics for this specific structure
                listTemp.Add(new Tuple<string, double, VolumePresentation, DoseValuePresentation>(itr.Item3, stat, vp, dvp));
            }
            //need one more add statement to ensure the final structure in the requested statistic list gets added
            statsRequest.Add(new Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>>(temp, temp2, listTemp));

            //retrieve the data from the plans and update the GUI with the requested information
            getStatsFromPlans();
            updateStats();
        }

        private void getStatsFromPlans()
        {
            //get stats from previous plans.
            //this function is relatively ugly because it is three nested loops. The first iterates through the structures in the requested statistics list, the second iterates through the requested statistics for 
            //a given structure, the third iterates through the plans to grab the requested statistic. This loop structure will grab the requested statistics for a given structure, one-by-one, until there are no more,
            //and it will move on to the next structure.
            foreach (Tuple<string, double, List<Tuple<string, double, VolumePresentation, DoseValuePresentation>>> itr in statsRequest)
            {
                //vector of structures. This is used to limit the number of times we have to grab the structure from each plan (very redundant if there are a lot of requested statistics for a given structure)
                List<Structure> structs = new List<Structure> { };
                //requested statistic, query value, units on requested statistic, vector to hold the retrieved statistic value from each plan
                List<Tuple<string, double, string, List<double>>> temp = new List<Tuple<string, double, string, List<double>>> { };
                foreach (Tuple<string, double, VolumePresentation, DoseValuePresentation> itr1 in itr.Item3)
                {
                    //initialize new list to hold the retrieved statistic from each plan
                    List<double> results = new List<double> { };
                    //current plan number in the list of plans
                    int count = 0;
                    //should the specific structure be retrieved from eahc plan? This should only be performed once per structure regardless of the number of requested statistics for that structure
                    bool getStructures = true;
                    //if this is not the first requested statistic for this structure, skip retrieving the structure from the plans
                    if (itr1 != itr.Item3.First()) getStructures = false;
                    foreach (BrachyPlanSetup p in plans)
                    {
                        if (getStructures)
                        {
                            //create an empty vector to hold all structures retrieved from the plans that match the id in the requested statistics array
                            List<Structure> SOI = new List<Structure> { };
                            if (!itr.Item1.Contains("ctv")) SOI = p.StructureSet.Structures.Where(x => x.Id.ToLower().Contains(itr.Item1.ToLower()) && !x.IsEmpty).ToList();
                            else
                            {
                                //special logic for the CTV structure because the CTV's will generally be labeled as ctv_mri<fraction number>
                                SOI = p.StructureSet.Structures.Where(x => x.Id.ToLower().Contains(String.Format("ctv_mri{0}", count+1)) && !x.IsEmpty).ToList();
                                //there will be cases where the ctv will NOT be contoured on the MRI, in which case there will be no ctv_mri structure in the structure set. In this case grab all structures that contain 'ctv'
                                if (SOI.Count == 0) SOI = p.StructureSet.Structures.Where(x => x.Id.ToLower().Contains("ctv") && !x.IsEmpty).ToList();
                            }
                            if (SOI.Count > 0)
                            {
                                //the retrieved list of structures is not empty
                                if (SOI.Count > 1)
                                {
                                    //the retrieved list of structures has more than one entry --> it is unclear which structure should be queried for the requested statistic. Ask the user which structure should be used
                                    autoSecondCheck.selectItem SUI = new autoSecondCheck.selectItem();
                                    SUI.title.Text = String.Format("Warning! Multiple non-empty structures \nwith string {0} found in plan {1}!\n\nPlease select a structure for evaluation!", itr.Item1, p.Id);
                                    foreach (Structure s in SOI) SUI.itemCombo.Items.Add(s.Id);
                                    SUI.itemCombo.Text = SOI.First().Id;
                                    SUI.ShowDialog();
                                    if (!SUI.confirm) return;
                                    structs.Add(p.StructureSet.Structures.First(x => x.Id == SUI.itemCombo.Text));
                                }
                                else structs.Add(SOI.First());
                                //if the dose is valid in the plan of interest, retrieve the requested statistic
                                if (p.IsDoseValid) results.Add(getData(p, structs.ElementAt(count), itr1));
                                else results.Add(0.0);
                            }
                            else
                            {
                                //the retrieved list of structures is empty
                                structs.Add(null);
                                MessageBox.Show(String.Format("Warning! No matching structure found for: {0} in plan: {1}! Skipping", itr.Item1, p.Id));
                                results.Add(0.0);
                            }
                        }
                        else
                        {
                            //structures have been previously retrieved from the plans. Now grab the requested statistic for that structure if it is not null
                            if (structs.ElementAt(count) == null)
                            {
                                MessageBox.Show(String.Format("Warning! No matching structure found for: {0} in plan: {1}! Skipping", itr.Item1, p.Id));
                                results.Add(0.0);
                            }
                            else
                            {
                                if (p.IsDoseValid) results.Add(getData(p, structs.ElementAt(count), itr1));
                                else results.Add(0.0);
                            }
                        }
                        count++;
                    }
                    //determine the string representation of the requested units on the requested statistic
                    string units = "%";
                    if (itr1.Item1.Contains("Dose at Volume") || itr1.Item1.Contains("Dmean")) { if (itr1.Item3 == VolumePresentation.AbsoluteCm3) units = "cc"; }
                    else if (itr1.Item4 == DoseValuePresentation.Absolute) units = "Gy";
                    //add the data for this requested statistic to the temp list, which will be added to the statsResults list after all requested statistics for the current structure have been retrieved
                    temp.Add(new Tuple<string, double, string, List<double>>(itr1.Item1, itr1.Item2, units, results));
                }
                //all requested statistics have been retrieved for the current structure --> add the data to the statsResults list (this is the aggregate data for this structure)
                statsResults.Add(new Tuple<string, double, List<Tuple<string, double, string, List<double>>>>(itr.Item1, itr.Item2, temp));
            }
        }

        //the function to actually grab the requested statistic for a given structure from a given plan
        private double getData(BrachyPlanSetup p, Structure s, Tuple<string, double, VolumePresentation, DoseValuePresentation> itr)
        {
            double value = 0.0;
            p.DoseValuePresentation = DoseValuePresentation.Absolute;
            if (itr.Item1.Contains("Dose at Volume"))
            {
                //a dose is requested. Syntax: structure, requested value, volume presentation, doseValuePresentation
                value = p.GetDoseAtVolume(s, itr.Item2, itr.Item3, itr.Item4).Dose;
                //if we want the retrieved statistic to have units of Gy, we need to divide by 100 (our Eclipse install is configured so the default units on dose are cGy)
                if (itr.Item4 == DoseValuePresentation.Absolute) value /= (100 * scaleFactor);
            }
            else if (itr.Item1.Contains("Volume at Dose"))
            {
                //a volume is requested
                DoseValue.DoseUnit d;
                double dose = itr.Item2;
                if (itr.Item4 == DoseValuePresentation.Absolute)
                {
                    d = DoseValue.DoseUnit.cGy;
                    //convert from Gy to cGy (Eclipse internal dose units for our install)
                    dose *= 100.0;
                }
                else d = DoseValue.DoseUnit.Percent;
                //syntax: structure, doseValue, volume representation
                value = p.GetVolumeAtDose(s, new DoseValue(dose, d), itr.Item3);
            }
            else if (itr.Item1.Contains("Dmean"))
            {
                //need to determine if the structure is pt A or an actual structure
                if (s.Id.ToLower().Contains("pt a"))
                {
                    //pt A is drawn as a reference line. Retrieve the two points that define the line and compute the mean dose from these two points
                    VVector[] points = s.GetReferenceLinePoints();
                    double meanPtADose = 0.0;
                    for (int i = 0; i < points.Count(); i++)
                    {
                        VVector loc = points.ElementAt(i);
                        meanPtADose += p.Dose.GetDoseToPoint(loc).Dose;
                    }
                    value = meanPtADose / points.Count();
                    //convert mean dose (in cGy) to either Gy or a percent
                    if (itr.Item4 == DoseValuePresentation.Absolute) value /= (100 * scaleFactor);
                    else value *= 100.0 / p.TotalDose.Dose;
                }
                else
                {
                    //it is an actual structure. Syntax: structure, doseValuePresentation, always relative volume presentation (doesn't matter), bin width in cGy
                    DVHData dvh = p.GetDVHCumulativeData(s, itr.Item4, VolumePresentation.Relative, 0.1);
                    if (dvh != null) value = dvh.MeanDose.Dose;
                    else value = 0.0;
                    if (itr.Item4 == DoseValuePresentation.Absolute) value /= (100 * scaleFactor);
                }
            }
            //special case if the user just wants the volume of the structure in the given plan
            else if (itr.Item1.Contains("Volume (cc)")) value = s.Volume;
            return value;
        }

        //function to update the text window with the retrieved results. Be careful messing with the formatting in the string.format functions
        private void updateStats()
        {
            results.Text = "";
            //variables to hold the EBRT total for the target and the OARs
            double tumorEBRTtotal = 0.0;
            double bladderEBRTtotal = 0.0, bowelEBRTtotal = 0.0, rectumEBRTtotal = 0.0, sigmoidEBRTtotal = 0.0;
            //plan, number of delivered fractions, tumor D2cc EQD2, bladder D2cc EQD2, bowel D2cc EQD2, rectum D2cc EQD2 (scaled for actual number of delivered fractions)
            List<Tuple<ExternalPlanSetup, int, double, double, double, double, double>> externalBeamResults = new List<Tuple<ExternalPlanSetup, int, double, double, double, double, double>> { };
            //external beam plan, number of treated fractions
            List<Tuple<ExternalPlanSetup, int>> ebPlans = new List<Tuple<ExternalPlanSetup, int>> { };
            //retrieve the previous external beam plans if we don't want to assume the max EQD2. This will always evalute to false per Dr. Kidd's request
            //if (!assumeMaxEQD2) ebPlans = getEBplans();
            if (ebPlans.Any())
            {
                if (ebPlans.Count() > 1)
                {
                    //multiple EBRT plans found. Set the EBRT dose per fraction and EBRT number of fraction textboxes to 1.00 and 1, respectively
                    EBRTdosePerFxTB.Text = String.Format("{0:0.00}", "1.00");
                    EBRTnumFxTB.Text = "1";
                }
                else
                {
                    //one EBRT plan found, update the EBRT dose per fractio nand number of fractions and update the relevant text boxes
                    p.EBRTdosePerFx = ebPlans.First().Item1.DosePerFraction.Dose / 100;
                    p.EBRTnumFx = (int)ebPlans.First().Item1.NumberOfFractions;
                    EBRTdosePerFxTB.Text = String.Format("{0:0.00}", p.EBRTdosePerFx);
                    EBRTnumFxTB.Text = p.EBRTnumFx.ToString();
                }

                foreach (Tuple<ExternalPlanSetup, int> itr in ebPlans)
                {
                    //retrieve the physical dose delivered from each external beam plan for the target and OAR structures. All we care about here is the D2cc for each of the structures
                    itr.Item1.DoseValuePresentation = DoseValuePresentation.Absolute;
                    if (itr.Item1.StructureSet.Structures.FirstOrDefault(x => x.Id.ToLower() == "bladder") != null) bladderEBRTtotal = itr.Item1.GetDoseAtVolume(itr.Item1.StructureSet.Structures.First(x => x.Id.ToLower() == "bladder"), 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose / 100;
                    if (itr.Item1.StructureSet.Structures.FirstOrDefault(x => x.Id.ToLower() == "bowel_bag") != null) bowelEBRTtotal = itr.Item1.GetDoseAtVolume(itr.Item1.StructureSet.Structures.First(x => x.Id.ToLower() == "bowel_bag"), 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose / 100;
                    if (itr.Item1.StructureSet.Structures.FirstOrDefault(x => x.Id.ToLower() == "rectum") != null) rectumEBRTtotal = itr.Item1.GetDoseAtVolume(itr.Item1.StructureSet.Structures.First(x => x.Id.ToLower() == "rectum"), 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose / 100;
                    if (itr.Item1.StructureSet.Structures.FirstOrDefault(x => x.Id.ToLower() == "sigmoid") != null) sigmoidEBRTtotal = itr.Item1.GetDoseAtVolume(itr.Item1.StructureSet.Structures.First(x => x.Id.ToLower() == "sigmoid"), 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose / 100;
                    tumorEBRTtotal = itr.Item1.GetDoseAtVolume(itr.Item1.StructureSet.Structures.First(x => x.Id == itr.Item1.TargetVolumeID), 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose / 100;
                    //EQD2 = (Dose from plan*(α/β + dose/fx)/(α/β + 2 Gy/fx))
                    //calculate tumor and OAR total EQD2's. Assume alpha/beta ratios of 10 Gy and 3 Gy for the tumor and OAR, respectively
                    externalBeamResults.Add(new Tuple<ExternalPlanSetup, int, double, double, double, double, double>(itr.Item1, itr.Item2,
                        tumorEBRTtotal * ((double)(itr.Item2) / (double)itr.Item1.NumberOfFractions) * ((10.0 + (itr.Item1.DosePerFraction.Dose / 100)) / (10.0 + 2.0)),
                        bladderEBRTtotal * ((double)(itr.Item2) / (double)itr.Item1.NumberOfFractions) * ((3.0 + (itr.Item1.DosePerFraction.Dose / 100)) / (3.0 + 2.0)),
                        bowelEBRTtotal * ((double)(itr.Item2) / (double)itr.Item1.NumberOfFractions) * ((3.0 + (itr.Item1.DosePerFraction.Dose / 100)) / (3.0 + 2.0)),
                        rectumEBRTtotal * ((double)(itr.Item2) / (double)itr.Item1.NumberOfFractions) * ((3.0 + (itr.Item1.DosePerFraction.Dose / 100)) / (3.0 + 2.0)),
                        sigmoidEBRTtotal * ((double)(itr.Item2) / (double)itr.Item1.NumberOfFractions) * ((3.0 + (itr.Item1.DosePerFraction.Dose / 100)) / (3.0 + 2.0))));
                }
            }
            else
            {
                //no external beam plan found or we want to assume the max EQD2. Assume alpha/beta ratios of 10 Gy and 3 Gy for the tumor and OAR, respectively
                tumorEBRTtotal = p.EBRTnumFx * p.EBRTdosePerFx * (p.EBRTdosePerFx + 10) / (2.0 + 10.0);
                bladderEBRTtotal = bowelEBRTtotal = rectumEBRTtotal = sigmoidEBRTtotal = p.EBRTnumFx * p.EBRTdosePerFx * (p.EBRTdosePerFx + 3.0) / (2.0 + 3.0);
            }
            //get the needle dwell time (i.e., not the Tandem, ovoids, VC, etc.) and the total dwell time for all applicators
            double needleDwellTime = getNeedleDwellTime(plan);
            double totalDwellTime = getDwellTime(plan.Catheters.ToList());

            //start populating the text window
            results.Text += "";
            //results.Document.Blocks.Add(new Paragraph(new Run("")));
            results.Text += String.Format("{0}", DateTime.Now.ToString()) + System.Environment.NewLine;
            //info about patient, id, plan, current fraction, etc.
            results.Text += String.Format("Patient Name: {0}, {1}", VMS.TPS.Script.GetScriptContext().Patient.LastName, VMS.TPS.Script.GetScriptContext().Patient.FirstName) + System.Environment.NewLine;
            results.Text += String.Format("Patient MRN: {0}", VMS.TPS.Script.GetScriptContext().Patient.Id) + System.Environment.NewLine;
            results.Text += String.Format("Current HDR fraction: {0}", plans.Count().ToString()) + System.Environment.NewLine;
            results.Text += String.Format("Total number of HDR fractions: {0}", numFractions) + System.Environment.NewLine;
            if (needleDwellTime > 0.0)
            {
                results.Text += String.Format("Number of needles: {0}", getNeedles(plan.Catheters.ToList(), plan).Count()) + System.Environment.NewLine;
                results.Text += String.Format("Needle relative dose contribution: {0:0.0}% ({1:0.0} of {2:0.0} seconds)", 100 * needleDwellTime / totalDwellTime, needleDwellTime, totalDwellTime) + System.Environment.NewLine + System.Environment.NewLine;
            }
            else
            {
                results.Text += String.Format("No needles present in current plan ({0})", plan.Id);
                results.Text += System.Environment.NewLine + System.Environment.NewLine;
            }
            results.Text += "------------------------------------------------------------------------------------------------------------------" + System.Environment.NewLine;

            //EQD2 data
            if (!ebPlans.Any())
            {
                //do not use external beam dose statistics for D2cc for the bowel, bladder, rectum, and tumor
                if (!assumeMaxEQD2) results.Text += String.Format(" NO EXTERNAL BEAM PLAN FOUND! ASSUMING MAX EQD2 FOR TUMOR AND ALL OARs!") + System.Environment.NewLine;
                results.Text += String.Format("{0,-34}", " EXTERNAL BEAM THERAPY:    ") + String.Format("{0,-19}", "Tumor    ") + String.Format("{0,-15}", "OAR") + System.Environment.NewLine;
                results.Text += String.Format("     {0,-24}     ", String.Format("Fx Dose (Gy): {0:0.00}", p.EBRTdosePerFx)) + String.Format("{0,-15}", "EQD2 [α/β=10Gy]    ") + String.Format("{0,-15}", "EQD2 [α/β=3Gy]") + System.Environment.NewLine;
                results.Text += String.Format("     {0,-24}     ", String.Format("Fx #: {0}", p.EBRTnumFx)) + String.Format("{0,-15:N1}    ", tumorEBRTtotal) + String.Format("{0,-15:N1}", bladderEBRTtotal) + System.Environment.NewLine;
            }
            else
            {
                //use external beam dose statistics for D2cc
                tumorEBRTtotal = bladderEBRTtotal = bowelEBRTtotal = rectumEBRTtotal = 0.0;
                results.Text += String.Format(" EXTERNAL BEAM PLAN DOSE CONTRIBUTIONS:") + System.Environment.NewLine;
                results.Text += String.Format("{0,-35}", "                            ") + String.Format("{0,-19}", "Tumor    ") + String.Format("{0,-19}", "bladder") + String.Format("{0,-19}", "bowel") + String.Format("{0,-19}", "rectum") + System.Environment.NewLine;
                results.Text += String.Format(" Plan Id    Fx Dose (Gy)  num Fx   ") + String.Format("{0,-15}", "EQD2 [α/β=10Gy]    ") + String.Format("{0,-15}", "EQD2 [α/β=3Gy]     ") + String.Format("{0,-15}", "EQD2 [α/β=3Gy]     ") + String.Format("{0,-15}", "EQD2 [α/β=3Gy]     ") + System.Environment.NewLine;
                foreach (Tuple<ExternalPlanSetup, int, double, double, double, double, double> itr in externalBeamResults)
                {
                    results.Text += String.Format(" {0,-13} {1:0.00}       {2,-2}       ", itr.Item1.Id, itr.Item1.DosePerFraction.Dose / 100, itr.Item2) + String.Format("{0,-15:N1}    ", itr.Item3) + String.Format("{0,-15:N1}    ", itr.Item4) + String.Format("{0,-15:N1}    ", itr.Item5) + String.Format("{0,-15:N1}    ", itr.Item6) + System.Environment.NewLine;
                    tumorEBRTtotal += itr.Item3;
                    bladderEBRTtotal += itr.Item4;
                    bowelEBRTtotal += itr.Item5;
                    rectumEBRTtotal += itr.Item6;
                    sigmoidEBRTtotal += itr.Item7;
                }
                if (ebPlans.Count() > 1) results.Text += String.Format("             TOTAL                 ") + String.Format("{0,-15:N1}    ", tumorEBRTtotal) + String.Format("{0,-15:N1}    ", bladderEBRTtotal) + String.Format("{0,-15:N1}    ", bowelEBRTtotal) + String.Format("{0,-15:N1}    ", rectumEBRTtotal) + System.Environment.NewLine;
            }
            results.Text += "------------------------------------------------------------------------------------------------------------------" + System.Environment.NewLine + System.Environment.NewLine;

            //format the header output for the retrieved statistics (fraction number, HDR sum, etc. columnwise)
            string message = "                               ";
            for (int i = 0; i < numFractions; i++) message += String.Format("| Fx #{0} ", i + 1);
            message += String.Format("| HDR sum | HDR+EBRT | Aim | Limit | Met? |") + System.Environment.NewLine;
            results.Text += message;

            //iterate through the retrieved statistics and add their results to the output window. The results will be listed for each structure where each structure is listed in alphabetical order
            //for any retrieved dose statistics reported in absolute units, the EQD2 will be calculated. 
            //structure id, alpha/beta, list of statistics, query value, units, vector of results from the plans for this structure
            helpers h = new helpers();
            foreach (Tuple<string, double, List<Tuple<string, double, string, List<double>>>> itr in statsResults)
            {
                //the way the text is added to the textblock changes for this section as special formatting is applied to the YES/NO in the 'met?' section of the line
                results.Inlines.Add(String.Format("{0}(α/β = {1}Gy)", itr.Item1, itr.Item2) + System.Environment.NewLine);
                //statistics, query value, units, vector of results from the plans
                foreach (Tuple<string, double, string, List<double>> itr1 in itr.Item3)
                {
                    //update this line of data. This is the physical data from each plan with no additional post-processing

                    //requested statistic = query value (units on query value)
                    if (itr1.Item1.Contains("Dose at Volume") || itr1.Item1.Contains("Volume at Dose")) results.Inlines.Add(String.Format(" {0,-30}", String.Format("{0} = {1:0.0}{2} ", itr1.Item1, itr1.Item2, itr1.Item3)));
                    //Dmean or Volume requested
                    else results.Inlines.Add(String.Format(" {0, -30}", itr1.Item1));

                    //iterate through the number of planned fractions and retrieve the statistics from the associated plans
                    for (int i = 0; i < numFractions; i++)
                    {
                        //dummy variable to hold the retrieved statistic
                        double value = 0.0;
                        //if the size of the results array (i.e., the number of plans) is greater than the current loop iteration count, use the value in the results vector
                        if (i < itr1.Item4.Count) value = itr1.Item4.ElementAt(i);
                        else
                        {
                            //otherwise propagate the value forward to all planned fractions ONLY IF THE REQUESTED STATISTIC IS A DOSE VALUE! If the requested statistic is a volume, report 0.0 (i.e., propagating a 
                            //volume forward makes no sense)
                            if (itr1.Item1.Contains("Dose at Volume") || itr1.Item1.Contains("Dmean")) value = itr1.Item4.Last();
                            else value = 0.0;
                        }
                        results.Inlines.Add(String.Format("| {0,-5:N1} ", value));

                        //if the requested statistic meets the criteria in the following if statements, then it is a metric of interest that should be saved so that it can be written to the excel file
                        int ind = p.excelStatistics.FindIndex(x => x.Item1 == itr.Item1 && x.Item2 == itr1.Item1 && x.Item3 == itr1.Item2 && x.Item4 == itr1.Item3);
                        if (ind != -1) excelData.ElementAt(ind).Add(value);
                        else if (itr1.Item1.Contains("Dmean") || itr1.Item1.Contains("Volume (cc)"))
                        {
                            ind = p.excelStatistics.FindIndex(x => x.Item1 == itr.Item1 && x.Item2 == itr1.Item1);
                            if (ind != -1) excelData.ElementAt(ind).Add(value);
                        }
                    }
                    results.Inlines.Add("|");
                    results.Inlines.Add(new LineBreak());

                    if (itr1.Item1 == "Dose at Volume (Gy)" || itr1.Item1 == "Dmean (Gy)")
                    {
                        //need to compute EQD2 values for the requested absolute doses
                        results.Inlines.Add(String.Format(" {0,-30}", String.Format("EQD2(α/β = {0}Gy) ", itr.Item2)));
                        double HDRsum = 0.0;
                        double HDR_EBRT_sum = 0.0;
                        double val = 0.0;
                        for (int i = 0; i < numFractions; i++)
                        {
                            //EQD2 is calculated for each structure for a SINGLE fraction (i.e., the current fraction). The single fraction EQD2 values are then added to obtain the cumulative EQD2. 
                            //If the current loop iteration is less than the size of the results vector, calculate EQD2 based on the value in the results vector, otherwise, propagate the last element in the vector
                            //forward to the remaining fractions
                            if (i < itr1.Item4.Count) val = itr1.Item4.ElementAt(i) * ((itr1.Item4.ElementAt(i) + itr.Item2) / (2.0 + itr.Item2));
                            else val = itr1.Item4.Last() * ((itr1.Item4.Last() + itr.Item2) / (2.0 + itr.Item2));
                            results.Inlines.Add(String.Format("| {0,-5:N1} ", val));
                            HDRsum += val;
                        }
                        //calculate the total EQD2 including HDR and EBRT
                        if (itr.Item1.Contains("gtv") || itr.Item1.Contains("ctv") || itr.Item1.Contains("pt A")) HDR_EBRT_sum = HDRsum + tumorEBRTtotal;
                        else
                        {
                            double total = 0.0;
                            //messy, but legacy code leftover from retrieving the EQD2 data from the actual external beam plans rather than assuming a particular dose was delivered
                            if (itr.Item1.Contains("bladder")) total = bladderEBRTtotal;
                            else if (itr.Item1.Contains("bowel")) total = bowelEBRTtotal;
                            else if (itr.Item1.Contains("rectum")) total = rectumEBRTtotal;
                            else if (itr.Item1.Contains("sigmoid")) total = sigmoidEBRTtotal;
                            HDR_EBRT_sum = HDRsum + total;
                        }
                        //report the HDR EQD2 and the EBRT+HDR EQD2
                        results.Inlines.Add(String.Format("| {0,-7:N1} | {1,-7:N1}  ", HDRsum, HDR_EBRT_sum));

                        //this is the logic to determine if the requested statistic has a dosimetric aim and/or limit that we are shooting for. Currently we have aims and limits for:
                        //Bladder D2cc, Bowel D2cc, Rectum D2cc, CTV D98%, CTV D90%, and PtA Dmean
                        //See the Gyn HDR BT spreadsheet for a list of current aims and limits
                        bool met = false;
                        Tuple<string, string> value = h.getAimLimit(p, itr.Item1, itr1.Item1, itr1.Item2, itr1.Item3);
                        string aim = value.Item1;
                        string limit = value.Item2;

                        //if either the aim or limit are nonempty, add these values to the reporting window text, close the final bracket, add the text to the window, and add a new line
                        if (aim != "" || limit != "")
                        {
                            met = h.checkIsMet(aim, limit, HDR_EBRT_sum);
                            results.Inlines.Add(new Run(String.Format("| {0,-3} | {1,-5} |", aim, limit)));
                            Run r = new Run(String.Format(" {2,-4} ", aim, limit, met ? "YES" : " NO")) { FontWeight = FontWeights.Bold };
                            if (met) r.Background = Brushes.LightGreen;
                            else r.Background = Brushes.LightPink;
                            results.Inlines.Add(r);
                            results.Inlines.Add("|");
                        }
                        else results.Inlines.Add("|");
                        results.Inlines.Add(new LineBreak());
                    }
                }
                //add a line of  ---- after each structure to make it look cleaner
                results.Inlines.Add("------------------------------------------------------------------------------------------------------------------");
                results.Inlines.Add(new LineBreak());
            }
            //results.Inlines.Add(new Run("test") { Background = Brushes.Red });
            resultsScroller.ScrollToBottom();
        }
        
        //get the dwell time of the needles only
        private double getNeedleDwellTime(BrachyPlanSetup p)
        {
            List<Catheter> needles = getNeedles(p.Catheters.ToList(), p);
            return getDwellTime(needles);
        }

        //get the needles used in the plan 
        private List<Catheter> getNeedles(List<Catheter> catheters, BrachyPlanSetup p)
        {
            //The needles are the catheters that are NOT the tandem, the ring, the ovoids, or the cylinder.
            List<Catheter> needles = catheters.Where(x => !x.Id.ToLower().Contains("tandem") && !x.Id.ToLower().Contains("ring") && !x.Id.ToLower().Contains("ovoid") && !x.Id.ToLower().Contains("cylinder")).ToList();
            needles.Sort(delegate(Catheter x, Catheter y) { return x.ChannelNumber.CompareTo(y.ChannelNumber); });
            //an extra piece of logic to account for certain cases where the patient is treated with a tandem and ring, but the treatment planner digitized the ring with a normal applicator (i.e., not a solid applicator with a predefined name)
            //and did not set the catheter Id to include 'ring'
            if (((p.ProtocolID.Contains("T&O") || p.ProtocolID.Contains("TO")) && (catheters.Count() - needles.Count() == 1)) || plan.SolidApplicators.Where(x => x.ApplicatorSetName.Contains("Universal Multi-channel Cylinder")).Any()) needles.RemoveAt(0);
            return needles;
        }

        //get the total dwell time of the catheters supplied as an arguement
        private double getDwellTime(List<Catheter> catheters)
        {
            double time = 0.0;
            foreach (Catheter c in catheters) time += c.GetTotalDwellTime();
            return time;
        }

        ////code used to query the aria database to grab the external beam plans for this patient. This function is not used per Dr. Kidd's request
        //private List<Tuple<ExternalPlanSetup, int>> getEBplans()
        //{
        //    List<Course> courses = VMS.TPS.Script.GetScriptContext().Patient.Courses.Where(x => !x.Id.ToLower().Contains("qa")).ToList();
        //    List<ExternalPlanSetup> plans = new List<ExternalPlanSetup> { };
        //    foreach (Course c in courses)
        //    {
        //        //List<ExternalPlanSetup> approvedPlans = c.ExternalPlanSetups.Where(x => x.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved).ToList();
        //        foreach (ExternalPlanSetup p in c.ExternalPlanSetups.Where(x => x.RTPrescription != null)) plans.Add(p);
        //    }
        //    //plan, number of delivered fractions
        //    List<Tuple<ExternalPlanSetup, int>> ebplans = new List<Tuple<ExternalPlanSetup, int>> { };

        //    try
        //    {
        //        using (Aria aria = new Aria())
        //        {
        //            ScriptContext context = VMS.TPS.Script.GetScriptContext();
        //            List<AriaQ_v15.Patient> pat = aria.Patients.Where(x => x.PatientId == context.Patient.Id).ToList();
        //            if (pat.Any())
        //            {

        //                long patientSer = pat.First().PatientSer;
        //                if (pat.Count() > 1) MessageBox.Show("patient number > 1");
        //                string message = "";
        //                foreach (ExternalPlanSetup p in plans)
        //                {
        //                    List<AriaQ_v15.Course> newCourses = aria.Courses.Where(tmp => tmp.PatientSer == patientSer && tmp.CourseId == p.Course.Id).ToList();
        //                    if (newCourses.Any())
        //                    {
        //                        long courseSer = newCourses.First().CourseSer;
        //                        //long planSetupSer = aria.PlanSetups.Where(tmp => (tmp.CourseSer == courseSer && tmp.PlanSetupId == p.Id)).First().PlanSetupSer;
        //                        if (aria.PlanSetups.FirstOrDefault(tmp => (tmp.CourseSer == courseSer && tmp.PlanSetupId == p.Id)) != null)
        //                        {
        //                            long planSetupSer = aria.PlanSetups.FirstOrDefault(tmp => (tmp.CourseSer == courseSer && tmp.PlanSetupId == p.Id)).PlanSetupSer;
        //                            //long planSetupSer = aria.PlanSetups.FirstOrDefault(tmp => tmp.PlanSetupId == p.Id).PlanSetupSer;
        //                            if (aria.RTPlans.FirstOrDefault(tmp => tmp.PlanSetupSer == planSetupSer) != null)
        //                            {
        //                                long RTPlanSer = aria.RTPlans.FirstOrDefault(tmp => tmp.PlanSetupSer == planSetupSer).RTPlanSer;
        //                                List<SessionRTPlan> sessionRTPlans = aria.SessionRTPlans.Where(tmp => tmp.RTPlanSer == RTPlanSer).ToList();
        //                                List<SessionProcedurePart> sessionProcedureParts = aria.SessionProcedureParts.Where(tmp => tmp.RTPlanSer == RTPlanSer).ToList();

        //                                message += String.Format("Plan: {0}", p.Id) + System.Environment.NewLine;
        //                                message += string.Format("num fx: {0}, dose per fx: {1:0.00} cGy", p.NumberOfFractions, p.DosePerFraction.Dose) + System.Environment.NewLine;
        //                                int count = 0;
        //                                foreach (SessionRTPlan srp in sessionRTPlans) if (srp.Status == "COMPLETE" || srp.Status == "TREAT") count++;
        //                                message += String.Format("Number of delivered fractions: {0}", count) + System.Environment.NewLine + System.Environment.NewLine;
        //                                if (count > 0) ebplans.Add(new Tuple<ExternalPlanSetup, int>(p, count));
        //                            }
        //                        }
        //                    }
        //                    else MessageBox.Show(string.Format("no course found for plan {0}", p.Id));
        //                }
        //                //display the plan id, number of fractions, dose per fraction, and number of completed fractions
        //                // MessageBox.Show(message);
        //            }
        //        }
        //    }
        //    catch (Exception e) { MessageBox.Show(e.Message); }

        //    return ebplans;
        //}

        //function to write the retrieved dose statistics to the excel spreadsheet for physician review. This function only writes the specified information to the spreadsheet (i.e., the relevant information is hard-coded into this function)
        //unfortunately, this function is messy, but necessary since there is not standard formatting for the excel spreadsheet for which statistic goes where (which makes it difficult to add/subtract metrics)
        private void WriteResultsExcel_Click(object sender, RoutedEventArgs e)
        {
            if (isIdealDoses) { MessageBox.Show("Error! Ideal doses are shown! Please uncheck the 'Show Ideal Doses' checkbox and try again!"); return; }
            // create Excel App
            Excel.Application myExcelApplication = new Excel.Application();
            Excel.Workbook myExcelWorkbook;
            Excel.Worksheet myExcelWorkSheet;
            myExcelApplication.DisplayAlerts = false; // turn off alerts

            // open the existing excel file
            if (!File.Exists(System.IO.Path.Combine(p.patientDataBase, p.excelTemplate))) { MessageBox.Show(string.Format("Error! The specified Excel file:\n{0}\ndoes not exist! Exiting", System.IO.Path.Combine(p.patientDataBase, p.excelTemplate))); return; }
            myExcelWorkbook = (Excel.Workbook)(myExcelApplication.Workbooks._Open(System.IO.Path.Combine(p.patientDataBase, p.excelTemplate),
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value));

            //get the first worksheet
            myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1];

            //write the requested header information to the excel file including patient name, EBRT num fx, EBRT dose per fx, etc.
            if (p.excelPatientName.Item1 != 0) myExcelWorkSheet.Cells[p.excelPatientName.Item1, p.excelPatientName.Item2] = String.Format("{0}, {1}", VMS.TPS.Script.GetScriptContext().Patient.LastName, VMS.TPS.Script.GetScriptContext().Patient.FirstName);
            if (p.excelPatientMRN.Item1 != 0) myExcelWorkSheet.Cells[p.excelPatientMRN.Item1, p.excelPatientMRN.Item2] = VMS.TPS.Script.GetScriptContext().Patient.Id;
            string physicianName = VMS.TPS.Script.GetScriptContext().Patient.PrimaryOncologistId;
            if (p.physicianIDs.Any()) p.physicianIDs.TryGetValue(physicianName, out physicianName);
            if (p.excelPhysician.Item1 != 0) myExcelWorkSheet.Cells[p.excelPhysician.Item1, p.excelPhysician.Item2] = physicianName;
            if (p.excelDate.Item1 != 0) myExcelWorkSheet.Cells[p.excelDate.Item1, p.excelDate.Item2] = DateTime.Now.ToString();
            if (p.excelTxSummary.Item1 != 0) myExcelWorkSheet.Cells[p.excelTxSummary.Item1, p.excelTxSummary.Item2] = String.Format("{0}Gy EBRT + HDR", p.EBRTdosePerFx * p.EBRTnumFx);
            if (p.excelEBRTdosePerFx.Item1 != 0) myExcelWorkSheet.Cells[p.excelEBRTdosePerFx.Item1, p.excelEBRTdosePerFx.Item2] = String.Format("{0:0.00}", p.EBRTdosePerFx);
            if (p.excelEBRTnumFx.Item1 != 0) myExcelWorkSheet.Cells[p.excelEBRTnumFx.Item1, p.excelEBRTnumFx.Item2] = String.Format("{0}", p.EBRTnumFx);
            if (p.excelEBRTtotalDose.Item1 != 0) myExcelWorkSheet.Cells[p.excelEBRTtotalDose.Item1, p.excelEBRTtotalDose.Item2] = String.Format("{0:0.00}", p.EBRTdosePerFx * p.EBRTnumFx);

            if (p.excelWriteFormat == "columnwise")
            {
                for (int i = 0; i < numFractions; i++)
                {
                    int ind = 0;
                    foreach (List<double> l in excelData)
                    {
                        char[] temp = p.excelStatistics.ElementAt(ind).Item6.ToCharArray();
                        for (int j = 0; j < i; j++) temp[0]++;
                        if (l.Any()) myExcelWorkSheet.Cells[p.excelStatistics.ElementAt(ind).Item5, temp[0].ToString()] = l.ElementAt(i);
                        ind++;
                    }
                }
                //add the number of needles and the needle dwell time to the spreadsheet
                if (p.excelNumNeedles.Item1 != 0)
                {
                    for (int i = 0; i < plans.Count; i++)
                    {
                        char[] temp = p.excelNeedleContr.Item2.ToCharArray();
                        for (int j = 0; j < i; j++) temp[0]++;
                        char[] temp2 = p.excelNumNeedles.Item2.ToCharArray();
                        for (int j = 0; j < i; j++) temp2[0]++;
                        if (getNeedles(plans.ElementAt(i).Catheters.ToList(), plans.ElementAt(i)).Count() > 0)
                        {
                            myExcelWorkSheet.Cells[p.excelNeedleContr.Item1, temp[0].ToString()] = 100 * getNeedleDwellTime(plans.ElementAt(i)) / getDwellTime(plans.ElementAt(i).Catheters.ToList());
                            myExcelWorkSheet.Cells[p.excelNumNeedles.Item1, temp2[0].ToString()] = getNeedles(plans.ElementAt(i).Catheters.ToList(), plans.ElementAt(i)).Count();
                        }
                        else
                        {
                            myExcelWorkSheet.Cells[p.excelNeedleContr.Item1, temp[0].ToString()] = 0.0;
                            myExcelWorkSheet.Cells[p.excelNumNeedles.Item1, temp2[0].ToString()] = 0;
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < numFractions; i++)
                {
                    int ind = 0;
                    foreach (List<double> l in excelData)
                    {
                        if (l.Any()) myExcelWorkSheet.Cells[p.excelStatistics.ElementAt(ind).Item5 + i, p.excelStatistics.ElementAt(ind).Item6] = l.ElementAt(i);
                        ind++;
                    }
                }
                //add the number of needles and the needle dwell time to the spreadsheet
                if (p.excelNumNeedles.Item1 != 0)
                {
                    for (int i = 0; i < plans.Count; i++)
                    {
                        if (getNeedles(plans.ElementAt(i).Catheters.ToList(), plans.ElementAt(i)).Count() > 0)
                        {
                            myExcelWorkSheet.Cells[p.excelNeedleContr.Item1 + i, p.excelNeedleContr.Item2] = 100 * getNeedleDwellTime(plans.ElementAt(i)) / getDwellTime(plans.ElementAt(i).Catheters.ToList());
                            myExcelWorkSheet.Cells[p.excelNumNeedles.Item1 + i, p.excelNumNeedles.Item2] = getNeedles(plans.ElementAt(i).Catheters.ToList(), plans.ElementAt(i)).Count();
                        }
                        else
                        {
                            myExcelWorkSheet.Cells[p.excelNeedleContr.Item1 + i, p.excelNeedleContr.Item2] = 0.0;
                            myExcelWorkSheet.Cells[p.excelNumNeedles.Item1 + i, p.excelNumNeedles.Item2] = 0;
                        }
                    }
                }
            }

            string result = new helpers().WriteResultsToExcel(p.patientDataBase, p.excelTemplate, myExcelWorkbook);
            if(result != "")
            {
                results.Inlines.Add(new LineBreak());
                results.Inlines.Add(result);
                results.Inlines.Add(new LineBreak());
                resultsScroller.ScrollToBottom();
            }
        }

        //simply drop the text in the reporting window to a text file
        private void WriteResultsText_Click(object sender, RoutedEventArgs e)
        {
            string fileName = new helpers().WriteResultsText(p.patientDataBase, results.Text);
            if (fileName != "")
            {
                results.Inlines.Add(new LineBreak());
                results.Inlines.Add(String.Format("Results written to txt file: {0}", fileName.Substring(fileName.LastIndexOf("\\") + 1, fileName.Length - fileName.LastIndexOf("\\") - 1)));
                results.Inlines.Add(new LineBreak());
                resultsScroller.ScrollToBottom();
            }
        }

        private void runSecondCheck_Click(object sender, RoutedEventArgs e)
        {
            doseStats.doseCalc calcWindow = new doseStats.doseCalc(plan, p.patientDataBase, p.useCurrentActivity, p.doseTolerance, secondCheckFile);
            calcWindow.ShowDialog();
        }
    }
}
