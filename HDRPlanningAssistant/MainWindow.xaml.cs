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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace HDRPlanningAssistant
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow(string[] args)
        {
            InitializeComponent();
            InitializeScript(args);
        }

        private void InitializeScript(string[] args)
        {
            throw new NotImplementedException();
        }

        #region help docs and shortcuts
        private void OpenHelpClick(object sender, RoutedEventArgs e)
        {

        }

        private void ShowShortcutsClick(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        private void EBRTdosePerFxTBTextChanged(object sender, TextChangedEventArgs e)
        {
            
        }
        private void EBRTnumFxTBTextChanged(object sender, TextChangedEventArgs e)
        {

        }

        #region request DVH metrics
        private void AddDVHMetricClick(object sender, RoutedEventArgs e)
        {

        }

        private void AddDefaultMetricsClick(object sender, RoutedEventArgs e)
        {

        }

        private void ClearResultsClick(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region Calculate metrics
        private void CalculateMetricsClick(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        private void ShowIdealDoses(object sender, RoutedEventArgs e)
        {

        }

        private void ShowMDAWindow(object sender, RoutedEventArgs e)
        {

        }

        #region write results
        private void WriteResultsTextClick(object sender, RoutedEventArgs e)
        {

        }

        private void WriteResultsExcelClick(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region second dose calc
        private void RunSecondCheckClick(object sender, RoutedEventArgs e)
        {

        }
        #endregion
    }
}
