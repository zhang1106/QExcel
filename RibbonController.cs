using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace QExcelAddIn
{
    [ComVisible(true)]
    public class RibbonController : Microsoft.Office.Core.IRibbonExtensibility
    {
        public const string DATA = "Data";
        public const string TESTDATA = "TestData";
        public const string RESULT = "Results";
        private DataAnalysisSvc _dataAnalysisSvc;
    
        public RibbonController()
        {
            _dataAnalysisSvc = new DataAnalysisSvc();
        }

        public string GetCustomUI(string ribbonID) =>
                @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                        <ribbon>
                           <tabs>
                                <tab id='test_tab' label='Quilt'>
                                    <group id='test_group' label='Test'>
                                        <button id='go' label='Go' size='large' onAction='OnGoClick'/>
                                    </group>
                                </tab>
                            </tabs>
                        </ribbon>
                    </customUI>";

        public void OnGoClick(Microsoft.Office.Core.IRibbonControl control)
        {
            if (ThisAddIn.CalcCount>0) {
                Debug.WriteLine("cant start another calculation.");
                return;
            };

            Interlocked.Increment(ref ThisAddIn.CalcCount);

            //start a timer
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            //create result sheet if not exists
            _dataAnalysisSvc.CreateSheet(RESULT);

            //start analysis
            Dictionary<string, int> fontCount = new Dictionary<string, int>();
            Dictionary<float, int> fontSizeCount = new Dictionary<float, int>();
            Dictionary<string, int> fillColorCount = new Dictionary<string, int>();
            List<int> list = new List<int>();

            Task t = Task.Factory.StartNew(() =>
                {
                    try
                    {
                        _dataAnalysisSvc.AnalyzeSheet(DATA, fontCount, fontSizeCount, fillColorCount, list);
                    }catch(Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
            ).ContinueWith((task) =>
            {
                try
                {
                    //calculate result
                    var mostFqtFont = fontCount.OrderByDescending(f => f.Value)
                        .Take(1).Single();
                    var mostFqtFontSize = fontSizeCount.OrderBy(f => f.Value)
                        .Take(1).Single();
                    var mostFqtFontColor = fillColorCount.OrderByDescending(f => f.Value)
                        .Take(1).Single();
                    var median = GetMedian(list);

                    string[] result = new[] {
                        $"Most Frequently Used Font: {mostFqtFont.Key} - {mostFqtFont.Value}",
                        $"Least Frequently Used Font Size: {mostFqtFontSize.Key} - {mostFqtFontSize.Value}",
                        $"Most Frequently Used Fill Color: {mostFqtFontColor.Key} - {mostFqtFontColor.Value}",
                        $"Median Number: {median} "
                    };

                    // write to results sheet
                    var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                    var resultSheet = (Excel.Worksheet)activeWorkbook.Worksheets[RESULT];

                    Excel.Range rslt = resultSheet.Range["A1:A5"];

                    for (int i = 0; i < 4; i++)
                    {
                        Debug.WriteLine(result[i]);
                        rslt.Cells[i + 1, 1].Value = result[i];
                    }

                    //stop timer
                    TimeSpan executionTime = stopwatch.Elapsed;
                    stopwatch.Stop();

                    rslt.Cells[5, 1].Value = $"Execution Time: {executionTime.TotalSeconds / 60.0} minutes";
                }
                catch(Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }
                finally
                {
                    Interlocked.Decrement(ref ThisAddIn.CalcCount);
                }
            });
        }

        private int GetMedian(List<int> list)
        {
            var sortedList = list.OrderBy(i => i).ToList();
            int mid = list.Count / 2;
            int median = 0;
            if (list.Count % 2 == 0)
            {
                median = (sortedList[mid - 1] + sortedList[mid]) / 2;
            }
            else
            {
                median = sortedList[mid];
            }
            return median;
        }
    }
}
