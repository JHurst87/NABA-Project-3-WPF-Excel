//Created by: Jordan Hurst
//Date: October 14, 2020
//Program takes in input from textboxes  when the Subumit button is clicked and saves them to an .xlsx database
//Program displays the .xlsx database file on the screen when the Print button is clicked.
//Program erases the database and displays a message if the file does not exist.
using System.Windows;
using DataController;
using System.IO;
using System.Data;
using System;

namespace PersonInfoExcel
{
    public partial class MainWindow : Window
    {
        readonly string path = @"C:\Users\thepa\Desktop\Internship\WPF\PersonInfoExcel\test.xlsx";
        public MainWindow()
        {
            InitializeComponent();
        }

        readonly DataControl data = new DataControl();
        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            string firstName = FirstName.Text;
            string lastName = LastName.Text;
            data.ToDatabase(firstName, lastName);
        }

        private void Print_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(path))
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
                Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                string strCellData = "";
                double douCellData;
                int rowCnt = 0;
                int colCnt = 0;

                DataTable dt = new DataTable();
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }

                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        try
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                            strData += strCellData + "|";
                        }
                        catch (Exception ex)
                        {
                            douCellData = (double)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    strData = strData.Remove(strData.Length - 1, 1);
                    dt.Rows.Add(strData.Split('|'));
                }

                dataGrid.ItemsSource = dt.DefaultView;

                excelBook.Close(true, null, null);
                excelApp.Quit();
            }

            else if (!File.Exists(path))
            {
                MessageBox.Show("File does not exist.");
            }
        }

        private void DeleteAll_Click(object sender, RoutedEventArgs e)
        {
            data.DeleteAllData();
        }
    }
}
