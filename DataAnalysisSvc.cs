using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Timers;

namespace QExcelAddIn
{
    public class DataAnalysisSvc
    {
        public void CreateSheet(string name)
        {
            // Get the active workbook
            var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            // Specify the name of the worksheet you want to create
            var worksheetName = name;

            // Check if the worksheet already exists in the workbook
            Excel.Worksheet worksheet = null;

            try
            {
                worksheet = (Excel.Worksheet)activeWorkbook.Worksheets[worksheetName];
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            if (worksheet == null)
            {
                // Create a new worksheet
                worksheet = (Excel.Worksheet)activeWorkbook.Worksheets.Add();
                worksheet.Name = worksheetName;
            }
        }

        public void AnalyzeSheet(string name,
            Dictionary<string, int> fontCount,
            Dictionary<float, int> fontSizeCount,
            Dictionary<string, int> fillColorCount,
            List<int> list)
        {
            // Get the active workbook
            var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            // Check if the worksheet already exists in the workbook
            Excel.Worksheet worksheet = null;

            try
            {
                worksheet = (Excel.Worksheet)activeWorkbook.Worksheets[name];
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            if(worksheet == null)
            {
                Debug.WriteLine($"{name} does not exist.");
                return;
            }

            // Find the last used cell in the worksheet
            Excel.Range lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

            // Get the range of all used cells
            Excel.Range usedRange = worksheet.Range["A1", lastCell];

            Debug.WriteLine($"Amout of used cells - {usedRange.Cells.Count}");
            Debug.WriteLine($"Started - 0.00/100");

            ProcessRange(usedRange, fontCount, fontSizeCount, fillColorCount, list);
        }

        public void ProcessRange(Excel.Range usedRange,
            Dictionary<string, int> fontCount,
            Dictionary<float, int> fontSizeCount,
            Dictionary<string, int> fillColorCount,
            List<int> list
            )
        {
            var count = 0;
            var total = usedRange.Cells.Count;

            //set a timer for reporting progress
            Timer timer = new Timer(3000);//3s
            timer.Elapsed += (sender, e) =>
            {
                Timer_Tick(sender, e, count, total);
            };
            timer.Start();

            //Access individual cells within the used range
            foreach (Excel.Range cell in usedRange)
            {
                if (cell.Font != null)
                {
                    //font name
                    if (!fontCount.ContainsKey(cell.Font.Name))
                    {
                        fontCount.Add(cell.Font.Name, 1);
                    }
                    else
                    {
                        fontCount[cell.Font.Name]++;
                    }

                    //font size
                    float fSize = (float)cell.Font.Size;
                    if (!fontSizeCount.ContainsKey(fSize))
                    {
                        fontSizeCount.Add(fSize, 1);
                    }
                    else
                    {
                        fontSizeCount[fSize]++;
                    }

                }

                //fill color
                Color fillColor = ColorTranslator.FromOle((int)cell.Interior.Color);
                string fillColorHex = fillColor.ToArgb().ToString("X6");

                if (!fillColorCount.ContainsKey(fillColorHex))
                {
                    fillColorCount.Add(fillColorHex, 1);
                }
                else
                {
                    fillColorCount[fillColorHex]++;
                }

                //assume int is enough
                try
                {
                    int cellValue = (int)cell.Value;

                    list.Add(cellValue);

                }catch(Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }

                count++;
            }

            timer.Stop();
            Timer_Tick(null, null, count, total);
        }

        // Timer tick event handler
        private void Timer_Tick(object sender, EventArgs e, int count, int total)
        {
            var progress = (count * 100.0) / total;
            Debug.WriteLine($"precessed {count} out of {total}, finished {progress:0.00}/100");
        }
    }
}
