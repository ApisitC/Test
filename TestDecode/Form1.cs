using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace TestDecode
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = "D:/LOGTEMP1.log";

            string excelPath = "D://Output.xlsx";

            // Open the file in binary mode
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (BinaryReader br = new BinaryReader(fs))
            {
                // Read the entire file into a byte array
                byte[] fileData = br.ReadBytes((int)fs.Length);

                // Create a new Excel workbook
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Calculate total rows needed
                int totalRows = (int)Math.Ceiling((double)fileData.Length / 16);

                // Write binary data to the worksheet
                int dataIndex = 0;
                for (int row = 1; row <= totalRows; row++)
                {
                    for (int col = 1; col <= 16; col++)
                    {
                        if (dataIndex < fileData.Length)
                        {
                            //int intValue = Convert.ToInt32(hexString, 16); // แปลงจากฐาน 16 เป็นเลขจำนวนเต็ม
                            //char charValue = Convert.ToChar(intValue); // แปลงจากเลขจำนวนเต็มเป็นตัวอักษร

                            worksheet.Cell(row, col).Value = fileData[dataIndex].ToString("X2");
                            dataIndex++;
                        }
                        else
                        {
                            // Fill remaining cells with empty string if data ends before completing a row
                            worksheet.Cell(row, col).Value = "";
                        }
                    }
                }

                // Save the workbook to a file
                workbook.SaveAs(excelPath);
            }

            Console.WriteLine("Data has been written to Excel file.");
        }
    }
}


//string filePath = "D:/LOGTEMP1.log";

//string excelPath = "D://Output.xlsx";
//// Open the file in binary mode
//            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
//            using (BinaryReader br = new BinaryReader(fs))
//            {
//                // Read the entire file into a byte array
//                byte[] fileData = br.ReadBytes((int)fs.Length);

//// Create a new Excel workbook
//var workbook = new XLWorkbook();
//var worksheet = workbook.Worksheets.Add("Sheet1");

//// Write binary data to the worksheet
//int row = 1;
//int col = 1;
//                foreach (byte b in fileData)
//                {
//                    worksheet.Cell(row, col).Value = b.ToString("X2");
//                    col++;
//                    if (col > 16)
//                    {
//                        col = 1;
//                        row++;
//                    }
//                }

//                // Save the workbook to a file
//                workbook.SaveAs(excelPath);
//            }

//            Console.WriteLine("Data has been written to Excel file.");
