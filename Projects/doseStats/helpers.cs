using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace doseStats
{
    class helpers
    {
        public helpers() 
        { }
        //get the patient folder if it exists in the GYN patient database directory
        public string getPatientFolder(string patientDataBase)
        {
            //grab all the folders and find the one that contains both the first and last name of the patient
            List<string> allFolders = Directory.GetDirectories(patientDataBase).ToList();
            List<string> newpath = allFolders.Where(x => x.ToLower().Contains(VMS.TPS.Script.GetScriptContext().Patient.LastName.ToLower()) && x.ToLower().Contains(VMS.TPS.Script.GetScriptContext().Patient.FirstName.ToLower())).ToList();

            //only one path was found (only one folder meets the above criteria)
            if (newpath.Count == 1) return newpath.First();

            //otherwise, no folder exists or there are multiple folders that meet this criteria. In this case, open a folder browser dialog box and request the user to select the appropriate folder
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.SelectedPath = patientDataBase;
            System.Windows.Forms.DialogResult result = fbd.ShowDialog();

            //some logic to ensure the selected folder is good and not the original patientDataBase directory
            if (result != System.Windows.Forms.DialogResult.OK && string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                MessageBox.Show("Path not found or path name NOT ok! Please try again!");
                return "";
            }
            if (string.Equals(patientDataBase.Substring(0, patientDataBase.Length - 1), fbd.SelectedPath))
            {
                MessageBox.Show("Please write the results to another directory!");
                return "";
            }
            return fbd.SelectedPath;
        }

        public string WriteResultsText(string patientDataBase, string message)
        {
            string fileName = "";
            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog
            {
                InitialDirectory = patientDataBase,
                Title = "Choose text file output",
                CheckPathExists = true,

                DefaultExt = "txt",
                Filter = "txt files (*.txt)|*.txt",
                FilterIndex = 2,
                RestoreDirectory = true,
            };

            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = saveFileDialog1.FileName;
                File.WriteAllText(fileName, message);
            }
            return fileName;
        }

        public string WriteResultsToExcel(string patientDataBase, string filename, Excel.Workbook myExcelWorkbook)
        {
            string result = "";

            //get the patient folder. If a bad folder path was returned, close the spreadsheet. Make the user try again
            string patientFolderPath = new helpers().getPatientFolder(patientDataBase);
            if (patientFolderPath == "")
            {
                myExcelWorkbook.Close(false);
                return "";
            }
            string filePath = patientFolderPath + @"\" + filename;

            try
            {
                // Save data in excel. No idea what most of these options do.
                myExcelWorkbook.SaveAs(filePath, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value);

                //close the worksheet
                myExcelWorkbook.Close(true, filePath, System.Reflection.Missing.Value);

                //if the excel spreadsheet was successfully saved, ask the user if they want to start excel now
                autoSecondCheck.confirmUI CUI = new autoSecondCheck.confirmUI();
                CUI.message.Text = "Results written to the Excel template." + Environment.NewLine + Environment.NewLine + "Start Excel?";
                CUI.button2.Text = "Yes";
                CUI.button1.Text = "No";
                CUI.ShowDialog();
                if (CUI.confirm) System.Diagnostics.Process.Start(filePath);
                result = System.Environment.NewLine + String.Format("Results written to excel file: {0}", filePath.Substring(filePath.LastIndexOf("\\") + 1, filePath.Length - filePath.LastIndexOf("\\") - 1)) + System.Environment.NewLine;
            }
            catch (Exception exception)
            {
                //something went wrong when trying to save the data. Likely causes are the folder couldn't be accessed, an excel file with the same name in that folder was open (so it couldn't be overwritted), etc.
                //In this case, the script will ask the user to specify a new name for the text file and/or write the file to another directory. This process will continue until the file is sucessfully written
                bool stillSucks = true;
                autoSecondCheck.confirmUI CUI = new autoSecondCheck.confirmUI();
                CUI.message.Text = String.Format("Error! Could not write results to excel template because:" + Environment.NewLine + "{0}" + Environment.NewLine + Environment.NewLine + "Change name of excel file?", exception.Message);
                CUI.button2.Text = "Yes";
                CUI.button1.Text = "No";
                CUI.ShowDialog();
                if (CUI.confirm)
                {
                    System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog
                    {
                        InitialDirectory = patientFolderPath,
                        FileName = filename,
                        CheckPathExists = true,

                        DefaultExt = ".xlsx",
                        Filter = "xlsx files (*.xlsx)|*.xlsx",
                        FilterIndex = 2,
                        RestoreDirectory = true,
                    };
                    while (stillSucks)
                    {
                        if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            try
                            {
                                //Save data in excel
                                myExcelWorkbook.SaveAs(saveFileDialog1.FileName, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                filePath = saveFileDialog1.FileName;

                                //close the worksheet
                                myExcelWorkbook.Close(true, saveFileDialog1.FileName, System.Reflection.Missing.Value);

                                //success!
                                stillSucks = false;
                                CUI = new autoSecondCheck.confirmUI();
                                CUI.message.Text = "Results written to the Excel." + Environment.NewLine + Environment.NewLine + "Start Excel?";
                                CUI.button2.Text = "Yes";
                                CUI.button1.Text = "No";
                                CUI.ShowDialog();
                                if (CUI.confirm) System.Diagnostics.Process.Start(saveFileDialog1.FileName);
                                result = System.Environment.NewLine + String.Format("Results written to excel file: {0}", filePath.Substring(filePath.LastIndexOf("\\") + 1, filePath.Length - filePath.LastIndexOf("\\") - 1)) + System.Environment.NewLine;
                            }
                            //something went wrong again. Reset the initial directory and excel file name and inform the user that they must try again
                            catch (Exception exception2) { saveFileDialog1.InitialDirectory = patientFolderPath; saveFileDialog1.FileName = result; MessageBox.Show(String.Format("NOPE: {0}. \nTRY AGAIN", exception2.Message)); }
                        }
                        else
                        {
                            //the case where the dialog results was NOT ok (i.e., the user hit the cancel button on the window)
                            stillSucks = false;
                            myExcelWorkbook.Close(false);
                        }
                    }
                }
                //the user does not want to write the results to another location. They would rather close the script, fix the problem, then try again.
                else myExcelWorkbook.Close(false);
            }

            return result;
        }
    }
}
