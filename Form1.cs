using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

               

                Excel.Application oXL = new Excel.Application();


#if DEBUG
                oXL.Visible = true;
                oXL.DisplayAlerts = false;
#else
                oXL.Visible = false; 
                oXL.DisplayAlerts = false;
#endif


                //Open the Excel File
                
                String inputFile = @"D:\Excel\InsertChart.xlsx";
                Excel.Workbook oWB = oXL.Workbooks.Open(inputFile);

                String SheetName = "ExperimentSheet";
                Excel._Worksheet oSheet = oWB.Sheets[SheetName];
                oSheet.Activate();
                //oSheet = oWB.ActiveSheet;
                //oSheet = oWB.Sheets.[1];

                //Insert an Empty Chart
                //Excel.Chart chart1 = oSheet.Ch .AddChart(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Excel.Shape chart1 = oSheet.Shapes.AddChart(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //Insert a basic 2-D bar graph
                chart1.Chart.ChartType = Excel.XlChartType.xlColumnClustered;
                
                //Set the Data Source for the chart
                Excel.Range Rng;
                Rng = oSheet.get_Range("A2", "C10");
                chart1.Chart.SetSourceData(Rng, Type.Missing);
                

                //Set the Chart Title
                chart1.Chart.HasTitle = true;
                chart1.Chart.ChartTitle.Text = "Population Growth";
               

                //Set the y-axis 
                var yaxis = (Excel.Axis)chart1.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                yaxis.HasTitle = true;
                yaxis.AxisTitle.Text = "Growth";
                yaxis.AxisTitle.Orientation = Excel.XlOrientation.xlVertical;


                //Set the X-axis 
                var xaxis = (Excel.Axis)chart1.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                xaxis.HasTitle = true;
                xaxis.AxisTitle.Text = "Population";
                yaxis.AxisTitle.Orientation = Excel.XlOrientation.xlHorizontal;


                //Set the Legend

                chart1.Chart.SeriesCollection(1).Name = (String)oSheet.get_Range("B1").Value2;
                chart1.Chart.SeriesCollection(2).Name = (String)oSheet.get_Range("C1").Value2;

                //Set the Size of the Chart
                chart1.Width = 350;
                chart1.Height = 350;
                chart1.Left = (float)oSheet.get_Range("D1").Left;
                chart1.Top = (float)oSheet.get_Range("D3").Top;

                oXL.Quit();

                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);

                oSheet = null;
                oWB = null;
                oXL = null;
                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
            }
            catch (Exception ex)
            {
                String errorMessage = "Error : " + ex.Message;
                MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        // Return the column name for this column number.
        private string ColumnNumberToName(int col_num)
        {
            // See if it's out of bounds.
            if (col_num < 1) return "A";

            // Calculate the letters.
            string result = "";
            while (col_num > 0)
            {
                // Get the least significant digit.
                col_num -= 1;
                int digit = col_num % 26;

                // Convert the digit into a letter.
                result = (char)((int)'A' + digit) + result;

                col_num = (int)(col_num / 26);
            }
            return result;
        }
    }
}
