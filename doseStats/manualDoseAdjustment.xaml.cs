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
using doseStats.Structs;

namespace doseStats
{
    /// <summary>
    /// Interaction logic for manualDoseAdjustment.xaml
    /// </summary>
    public partial class manualDoseAdjustment : Window
    {
        RoutedCommand closeWindowMacro = new RoutedCommand();
        //number of HDR fractions
        int numfx = 0;
        //data memebers to hold the retrieved statistics
        List<Tuple<string, double, List<Tuple<string, double, string, List<double>>>>> stats = new List<Tuple<string, double, List<Tuple<string, double, string, List<double>>>>> { };
        //a simplified list of the retrieved statistics so there is one row for each statistic (even if there are two statistics for the same structure)
        List<Tuple<string, double, string, double, string, List<double>>> simpleStatsList = new List<Tuple<string, double, string, double, string, List<double>>> { };
        Parameters p;
        double tumorEBRTtotal = 0.0;
        double bladderEBRTtotal = 0.0, bowelEBRTtotal = 0.0, rectumEBRTtotal = 0.0, sigmoidEBRTtotal = 0.0;

        public manualDoseAdjustment(int n, List<Tuple<string, double, List<Tuple<string, double, string, List<double>>>>> data, Parameters config)
        {
            InitializeComponent();
            //add a Ctrl+Q macro to make it easy to quit this window
            closeWindowMacro.InputGestures.Add(new KeyGesture(Key.Q, ModifierKeys.Control));
            CommandBindings.Add(new CommandBinding(closeWindowMacro, closeWindow));
            numfx = n;
            stats = new List<Tuple<string, double, List<Tuple<string, double, string, List<double>>>>>(data);
            p = config;
            add_header();
            addData();
        }

        private void closeWindow(object sender, RoutedEventArgs e) { this.Close(); }

        private void add_header()
        {
            //Structure, statistic, fx doses, EQD2 dose (EBRT + HDR), isMet? 
            StackPanel sp1 = new StackPanel();
            sp1.Height = 30;
            sp1.Width = statsSP.Width;
            sp1.Orientation = Orientation.Horizontal;
            sp1.Margin = new Thickness(5, 0, 5, 5);

            Label strName = new Label();
            strName.Content = "Structure";
            strName.HorizontalAlignment = HorizontalAlignment.Center;
            strName.VerticalAlignment = VerticalAlignment.Top;
            strName.Width = 80;
            strName.FontSize = 14;
            strName.Margin = new Thickness(0, 0, 0, 0);
            sp1.Children.Add(strName);

            Label statName = new Label();
            statName.Content = "Statistic";
            statName.HorizontalAlignment = HorizontalAlignment.Center;
            statName.VerticalAlignment = VerticalAlignment.Top;
            statName.Width = 80;
            statName.FontSize = 14;
            statName.Margin = new Thickness(0, 0, 0, 0);
            sp1.Children.Add(statName);

            //each item in stats should have the same number of dose entries (i.e., HDR fractions)
            for (int i = 0; i < stats.First().Item3.First().Item4.Count(); i++)
            {
                Label fxDoseLabel = new Label();
                fxDoseLabel.Content = String.Format("Fx {0} (Gy)",i+1);
                fxDoseLabel.HorizontalAlignment = HorizontalAlignment.Center;
                fxDoseLabel.VerticalAlignment = VerticalAlignment.Top;
                fxDoseLabel.Width = 80;
                fxDoseLabel.FontSize = 14;
                fxDoseLabel.Margin = new Thickness(0, 0, 0, 0);
                sp1.Children.Add(fxDoseLabel);
            }

            Label eqd2Label = new Label();
            eqd2Label.Content = "EQD2";
            eqd2Label.HorizontalAlignment = HorizontalAlignment.Center;
            eqd2Label.VerticalAlignment = VerticalAlignment.Top;
            eqd2Label.Width = 60;
            eqd2Label.FontSize = 14;
            eqd2Label.Margin = new Thickness(0, 0, 5, 0);
            sp1.Children.Add(eqd2Label);

            Label metLabel = new Label();
            metLabel.Content = "Met?";
            metLabel.HorizontalAlignment = HorizontalAlignment.Center;
            metLabel.VerticalAlignment = VerticalAlignment.Top;
            metLabel.Width = 60;
            metLabel.FontSize = 14;
            metLabel.Margin = new Thickness(0, 0, 0, 0);
            sp1.Children.Add(metLabel);

            statsSP.Children.Add(sp1);
        }

        private void addData()
        {
            tumorEBRTtotal = p.EBRTnumFx * p.EBRTdosePerFx * (p.EBRTdosePerFx + 10.0) / (2.0 + 10.0);
            bladderEBRTtotal = bowelEBRTtotal = rectumEBRTtotal = sigmoidEBRTtotal = p.EBRTnumFx * p.EBRTdosePerFx * (p.EBRTdosePerFx + 3.0) / (2.0 + 3.0);
            helpers h = new helpers();
            foreach (Tuple<string, double, List<Tuple<string, double, string, List<double>>>> itr in stats)
            {
                //statistics, query value, units, vector of results from the plans
                foreach (Tuple<string, double, string, List<double>> itr1 in itr.Item3)
                {
                    if (itr1.Item1 == "Dose at Volume (Gy)" || itr1.Item1 == "Dmean (Gy)")
                    {
                        //need to compute EQD2 values for the requested absolute doses
                        double HDRsum = 0.0;
                        double HDR_EBRT_sum = 0.0;
                        for (int i = 0; i < numfx; i++)
                        {
                            //EQD2 is calculated for each structure for a SINGLE fraction (i.e., the current fraction). The single fraction EQD2 values are then added to obtain the cumulative EQD2. 
                            //If the current loop iteration is less than the size of the results vector, calculate EQD2 based on the value in the results vector, otherwise, propagate the last element in the vector
                            //forward to the remaining fractions
                            if (i < itr1.Item4.Count) HDRsum += itr1.Item4.ElementAt(i) * ((itr1.Item4.ElementAt(i) + itr.Item2) / (2.0 + itr.Item2)); 
                            else HDRsum += itr1.Item4.Last() * ((itr1.Item4.Last() + itr.Item2) / (2 + itr.Item2));
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

                            StackPanel sp = new StackPanel();
                            sp.Height = 30;
                            sp.Width = statsSP.Width;
                            sp.Orientation = Orientation.Horizontal;
                            sp.Margin = new Thickness(5);

                            Label strName = new Label();
                            strName.Content = itr.Item1;
                            strName.Name = "Structure";
                            strName.HorizontalAlignment = HorizontalAlignment.Center;
                            strName.VerticalAlignment = VerticalAlignment.Top;
                            strName.Width = 80;
                            strName.FontSize = 14;
                            strName.Margin = new Thickness(0, 0, 0, 0);
                            sp.Children.Add(strName);

                            Label statType = new Label();
                            statType.Content = String.Format("{0,-30}", String.Format("D{0:0.0}{1} ", itr1.Item2, itr1.Item3));
                            statType.HorizontalAlignment = HorizontalAlignment.Center;
                            statType.VerticalAlignment = VerticalAlignment.Top;
                            statType.Width = 80;
                            statType.FontSize = 14;
                            statType.Margin = new Thickness(0, 0, 0, 0);
                            sp.Children.Add(statType);

                            for (int i = 0; i < itr1.Item4.Count; i++)
                            {
                                TextBox dose_tb = new TextBox();
                                dose_tb.Name = "dose_tb";
                                dose_tb.Width = 50;
                                dose_tb.Height = sp.Height - 5;
                                dose_tb.HorizontalAlignment = HorizontalAlignment.Left;
                                dose_tb.VerticalAlignment = VerticalAlignment.Center;
                                dose_tb.VerticalContentAlignment = VerticalAlignment.Center;
                                dose_tb.Margin = new Thickness(5, 0, 25, 0);
                                dose_tb.Text = String.Format("{0:0.00}", itr1.Item4.ElementAt(i));
                                dose_tb.TextAlignment = TextAlignment.Center;
                                dose_tb.BorderBrush = Brushes.Black;
                                dose_tb.BorderThickness = new Thickness(1.2);
                                if(i != itr1.Item4.Count - 1)
                                {
                                    dose_tb.IsReadOnly = true;
                                    dose_tb.Focusable = false;
                                    dose_tb.IsTabStop = false;
                                    dose_tb.Background = Brushes.LightGray;
                                }
                                dose_tb.GotFocus += Dose_tb_GotFocus;
                                dose_tb.TextChanged += Dose_tb_TextChanged;
                                sp.Children.Add(dose_tb);
                            }

                            TextBox eqd2_tb = new TextBox();
                            eqd2_tb.Name = "eqd2_tb";
                            eqd2_tb.Width = 50;
                            eqd2_tb.Height = sp.Height - 5;
                            eqd2_tb.HorizontalAlignment = HorizontalAlignment.Left;
                            eqd2_tb.VerticalAlignment = VerticalAlignment.Center;
                            eqd2_tb.VerticalContentAlignment = VerticalAlignment.Center;
                            eqd2_tb.Margin = new Thickness(0, 0, 10, 0);
                            eqd2_tb.Text = String.Format("{0:0.00}", HDR_EBRT_sum);
                            eqd2_tb.TextAlignment = TextAlignment.Center;
                            eqd2_tb.IsReadOnly = true;
                            eqd2_tb.Background = Brushes.LightGray;
                            eqd2_tb.BorderBrush = Brushes.Black;
                            eqd2_tb.BorderThickness = new Thickness(1.2);
                            eqd2_tb.Focusable = false;
                            eqd2_tb.IsTabStop = false;
                            sp.Children.Add(eqd2_tb);

                            TextBox ismet_tb = new TextBox();
                            ismet_tb.Name = "ismet_tb";
                            ismet_tb.Width = 50;
                            ismet_tb.Height = sp.Height - 5;
                            ismet_tb.HorizontalAlignment = HorizontalAlignment.Left;
                            ismet_tb.VerticalAlignment = VerticalAlignment.Center;
                            ismet_tb.VerticalContentAlignment = VerticalAlignment.Center;
                            ismet_tb.Margin = new Thickness(0, 0, 0, 0);
                            ismet_tb.TextAlignment = TextAlignment.Center;
                            ismet_tb.IsReadOnly = true;
                            ismet_tb.FontWeight = FontWeights.Bold;
                            if (met) { ismet_tb.Text = "YES"; ismet_tb.Background = Brushes.ForestGreen; }
                            else { ismet_tb.Text = "NO"; ismet_tb.Background = Brushes.Red; }
                            ismet_tb.BorderBrush = Brushes.Black;
                            ismet_tb.BorderThickness = new Thickness(1.2);
                            ismet_tb.Focusable = false;
                            ismet_tb.IsTabStop = false;
                            sp.Children.Add(ismet_tb);

                            statsSP.Children.Add(sp);
                            //for each row added to statsSP, copy the appropriate statistic and data to the simpleStatsList 
                            simpleStatsList.Add(new Tuple<string,double,string,double,string,List<double>>(itr.Item1, itr.Item2, itr1.Item1, itr1.Item2, itr1.Item3, new List<double>(itr1.Item4)));
                        }
                    }
                }
            }
        }

        //when a textbox is selected, highlight all of the text (makes it a bit easier to navigate)
        private void Dose_tb_GotFocus(object sender, RoutedEventArgs e) { (sender as TextBox).SelectAll(); }

        private void Dose_tb_TextChanged(object sender, TextChangedEventArgs e)
        {
            //data we will need from this function including the textbox where the text was modified, the eqd2 textbox, which row of children in statsSP contains the textbox, bools to keep track of which child element we are at when
            //iterating though the children in the stack panel
            TextBox tb = (TextBox)sender;
            TextBox eqd2 = null;
            bool row = false;
            bool doubleRow = false;
            int rowInd = 0;
            double hdrDose = 0.0;
            double.TryParse(tb.Text, out hdrDose);

            foreach (object obj in statsSP.Children)
            {
                foreach (object obj1 in ((StackPanel)obj).Children)
                {
                    if (row)
                    {
                        //save the eqd2 textbox as a variable. We want to iterate to the next child element (should be the isMet textbox) and perform operations on the current hdr dose textbox, eqd2 textbox, and isMet textbox.
                        eqd2 = obj1 as TextBox;
                        doubleRow = true;
                        row = false;
                    }
                    else if (doubleRow)
                    {
                        //need to substract 1 from rowInd to account for the header object stackpanel!
                        //update the EQD2 and isMet text boxes
                        updateEQD2Text(eqd2, hdrDose, rowInd - 1, (TextBox)obj1);
                        return;
                    }
                    //the textbox has a unique tag to it, so we can just loop through all children in the statSP children list and find which TextBox is equivalent to our Textbox. Once we find the right
                    //textbox, set the row flag to true. That way we can perform operations on the next child element (should be the eqd2 textbox).
                    if (obj1.Equals(tb)) row = true;
                }
                rowInd++;
            }
        }

        private void updateEQD2Text(TextBox eqd2, double dose, int rowInd, TextBox isMet)
        {
            double updateEQD2 = 0.0;
            //get the appropriate structure/statistic data based on the modified text box
            Tuple<string, double, string, double, string, List<double>> valStat = simpleStatsList.ElementAt(rowInd);
            //MessageBox.Show(String.Format("{0}, {1}, {2}, {3}, {4}", valStat.Item1, valStat.Item2, valStat.Item3, valStat.Item4, valStat.Item5));
            //first get the EBRT EQD2 dose contribution
            if (valStat.Item1 == "ctv" || valStat.Item1 == "gtv") updateEQD2 = tumorEBRTtotal;
            else if (valStat.Item1.Contains("bladder")) updateEQD2 = bladderEBRTtotal;
            else if (valStat.Item1.Contains("bowel")) updateEQD2 = bowelEBRTtotal;
            else if (valStat.Item1.Contains("rectum")) updateEQD2 = rectumEBRTtotal;
            else if (valStat.Item1.Contains("sigmoid")) updateEQD2 = sigmoidEBRTtotal;

            //next, sum up the contributions from the HDR fractions. Be sure NOT to include the current fraction dose (i.e., the achied dose in the current fraction)! Use the dose entered into the textbox instead!
            for(int i = 0; i < numfx; i++)
            {
                if (i < valStat.Item6.Count - 1) updateEQD2 += (valStat.Item6.ElementAt(i) * ((valStat.Item6.ElementAt(i) + valStat.Item2) / (2 + valStat.Item2))); 
                else updateEQD2 += (dose * ((dose + valStat.Item2) / (2 + valStat.Item2)));
            }
            //update the EQD2 text
            eqd2.Text = String.Format("{0:0.00}", updateEQD2);
            
            //get the aims/limits for this particular structure/statistic combination
            helpers h = new helpers();
            Tuple<string, string> value = h.getAimLimit(p, valStat.Item1, valStat.Item3, valStat.Item4, valStat.Item5);
            string aim = value.Item1;
            string limit = value.Item2;

            //update the formatting of the isMet textbox depending on if the constraint was met
            if (h.checkIsMet(aim, limit, updateEQD2))
            {
                isMet.Text = "YES"; 
                isMet.Background = Brushes.ForestGreen;
            }
            else
            {
                isMet.Text = "NO"; 
                isMet.Background = Brushes.Red;
            }
        }
    }
}
