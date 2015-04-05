using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AverageDoseValues
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }
        String filePath = "";
        public static int curMonth = 0;
        public static int curYear = 0;
        public static bool initDateFlag = true;
        List<decimal> listOfAvgDoseValues = new List<decimal>();
        List<string> listOfdateValues = new List<string>();
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;


        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                System.IO.StreamReader sr = new
                   System.IO.StreamReader(openFileDialog1.FileName);
            }
            filePath = openFileDialog1.FileName;
            textBox1.Text = filePath; // getting the file name from user input.
        }

        private void button1_Click(object sender, EventArgs e)
        {
             Excel.Range range;
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(filePath); //Giving the path to excel workbook to open and read the file.
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here

            string str = "";
            int rCnt, cCnt, dateCol = 0, doseCol = 0;
            range = MySheet.UsedRange; //range contains all the cells taht are currently in use.
            decimal avgOfDose = 0;
            int count = 0;
            //loop to iterate through the columns which contain Total and  Date, which we use to find the average values!
            //Cursor points to the required fields.
            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                str = Convert.ToString((range.Cells[1, cCnt] as Excel.Range).Value2);
                if (!String.IsNullOrEmpty(str))
                {
                    if (str.Contains("Date")) // Storing the column number in dateCol which has Date in the first row
                    {
                        dateCol = cCnt;

                    }
                    if (str.Contains("Total")) // Storing the column number in doseCol which has Total in the first row.
                    {

                        doseCol = cCnt;
                    }
                }
            }
            //loop to iterate through the rows that are taken from the first loop, this takes all the date values and total dose values in the selected file. 
            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {

                DateTime dateOfTEPC = new DateTime();
                decimal doseValue = 0;
                bool isValidTEPCDate = false; // Boolean to check for valid TEPC date
                bool isValidDoseValue = false; // Boolean to check for valid Dose Value


                object inpDate = (range.Cells[rCnt, dateCol] as Excel.Range).Value2; // gets the range of cells, rCnt is the rown count which contains the current row number of the particular column number of date
                if (inpDate != null)
                {
                    if (inpDate is double) // check if date is double
                    {
                        dateOfTEPC = DateTime.FromOADate((double)inpDate); // conver from OADate to Date data type
                        isValidTEPCDate = true;
                    }
                    else
                    {
                        DateTime.TryParse((string)inpDate, out dateOfTEPC); //  Parsing the input date to DateTime
                        isValidTEPCDate = true;
                    }


                }


                object dValue = (range.Cells[rCnt, doseCol] as Excel.Range).Value2; // Gets the range of cells which contain dose values, rCnt is the variable which iterated through all the dose values  of the aprticular row.
                if (dValue != null) //check for dose value for null. the cell in the excel may be empty
                {
                    if (dValue is double)
                    {
                        doseValue = Convert.ToDecimal(dValue);
                        isValidDoseValue = true;
                    }
                }

                if (isValidDoseValue && isValidTEPCDate) // Check if both Dose and TEPCDate are valid
                {

                    if (initDateFlag) // Initial date flag to check for the cursor reaches another month row.
                    {
                        curMonth = dateOfTEPC.Month;
                        curYear = dateOfTEPC.Year;
                        listOfdateValues.Add(dateOfTEPC.ToOADate().ToString()); // Adding corresponding date to listOdateValues list.
                        initDateFlag = false; // Update the boolean 
                    }


                    if (dateOfTEPC.Month == curMonth && dateOfTEPC.Year == curYear) // Incrementing the count (used to divide sum of dose values with) 
                    //and adding the dose values by checking current month and year
                    {
                        count++;
                        avgOfDose += doseValue;

                    }
                    else
                    {
                        curMonth = dateOfTEPC.Month;
                        curYear = dateOfTEPC.Year;
                        avgOfDose = avgOfDose / count; // calculating the average dose value
                        listOfdateValues.Add(dateOfTEPC.ToOADate().ToString());
                        // listOfYearValues.Add(curYear);
                        listOfAvgDoseValues.Add(avgOfDose);
                        count = 1; //Updating the count value for further iterations of for loop
                        //  MessageBox.Show(avgOfDose.ToString());
                        avgOfDose = doseValue;
                    }
                }

            }
            listOfAvgDoseValues.Add(avgOfDose / count); // Finally adding the average dose values to the list



            MessageBox.Show("Read Successfull");
            try
            {
                MyBook.Close(true, filePath, null);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to close spreadsheet");
            }


            System.IO.StreamWriter sw = new System.IO.StreamWriter(System.IO.Directory.GetCurrentDirectory()+"/AverageDoses.csv");

            for (int ii = 0; ii < listOfAvgDoseValues.Count; ii++)
            {

                sw.WriteLine(listOfdateValues[ii] + " ,  " + listOfAvgDoseValues[ii]); // Writing the obtained Avg Dose values into a CSV file with respect to the corresponding dates.

            }

            sw.Close();
            MessageBox.Show("Write Successful");

        
        }
    }
}
