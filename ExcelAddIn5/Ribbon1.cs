using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Diagnostics;
using System.Text;
using System.Web.Script.Serialization;

namespace ExcelAddIn5
{
    public partial class Ribbon1
    { 
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        /************  TRAIN BUTTON  ************/

        private void Train_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook actbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            Excel.Worksheet InputSheet = actbook.Sheets[1];

            Excel.Worksheet DataSheet = actbook.Sheets[2];

            Excel.Worksheet IndicatorSheet = actbook.Sheets[3];

            Excel.Worksheet OutputSheet = actbook.Sheets[4];

            Excel.Worksheet ErrorSheet = actbook.Sheets[6];

            double size = InputSheet.Cells[7, 2].Value2;

            Excel.Range DataRange1 = DataSheet.Range[DataSheet.Cells[2, 1], DataSheet.Cells[(int)size + 1, 1]];

            Excel.Range DataRange2 = DataSheet.Range[DataSheet.Cells[2, 2], DataSheet.Cells[(int)size + 1, 2]];

            Excel.Range DataRange3 = DataSheet.Range[DataSheet.Cells[2, 3], DataSheet.Cells[(int)size + 1, 3]];

            Excel.Range DataRange4 = DataSheet.Range[DataSheet.Cells[2, 4], DataSheet.Cells[(int)size + 1, 4]];

            Excel.Range DataRange5 = DataSheet.Range[DataSheet.Cells[2, 5], DataSheet.Cells[(int)size + 1, 5]];

            Excel.Range DataRange6 = DataSheet.Range[DataSheet.Cells[2, 6], DataSheet.Cells[(int)size + 1, 6]];

            Excel.Range OutRange1 = OutputSheet.Range[OutputSheet.Cells[2, 1], OutputSheet.Cells[(int)size + 1, 1]];

            Excel.Range OutRange2 = OutputSheet.Range[OutputSheet.Cells[2, 2], OutputSheet.Cells[(int)size + 1, 2]];

            Excel.Range OutRange3 = OutputSheet.Range[OutputSheet.Cells[2, 3], OutputSheet.Cells[(int)size + 1, 3]];

            Excel.Range OutRange4 = OutputSheet.Range[OutputSheet.Cells[2, 4], OutputSheet.Cells[(int)size + 1, 4]];

            Excel.Range OutRange5 = OutputSheet.Range[OutputSheet.Cells[2, 5], OutputSheet.Cells[(int)size + 1, 5]];

            Excel.Range OutRange6 = OutputSheet.Range[OutputSheet.Cells[2, 6], OutputSheet.Cells[(int)size + 1, 6]];

            Excel.Range OutRange7 = OutputSheet.Range[OutputSheet.Cells[1, 7], OutputSheet.Cells[(int)size + 1, 7]];

            Excel.Range ErrorRange1 = ErrorSheet.Range[ErrorSheet.Cells[2, 1], ErrorSheet.Cells[(int)size + 1, 1]];

            Excel.Range ErrorRange2 = ErrorSheet.Range[ErrorSheet.Cells[2, 2], ErrorSheet.Cells[(int)size + 1, 2]];

            Excel.Range ErrorRange3 = ErrorSheet.Range[ErrorSheet.Cells[2, 3], ErrorSheet.Cells[(int)size + 1, 3]];

            Excel.Range ErrorRange4 = ErrorSheet.Range[ErrorSheet.Cells[2, 4], ErrorSheet.Cells[(int)size + 1, 4]];

            Excel.Range ErrorRange5 = ErrorSheet.Range[ErrorSheet.Cells[2, 5], ErrorSheet.Cells[(int)size + 1, 5]];

            Excel.Range ErrorRange6 = ErrorSheet.Range[ErrorSheet.Cells[2, 6], ErrorSheet.Cells[(int)size + 1, 6]];

            const string open = "Open";
            const string high = "High";
            const string low = "Low";
            const string close = "Close";
            /****  Indicator Variables  ****/
            const string AC = "AC";
            const string AD = "AD";
            const string ADX = "ADX";
            const string Alligator = "Alligator";
            const string AO = "AO";
            const string ATR = "ATR";
            const string BearsPower = "BearsPower";
            const string Bands = "Bands";
            const string BullsPower = "BullsPower";
            const string CCI = "CCI";
            const string Custom = "Custom";
            const string DeMarker = "DeMarker";
            const string Envelopes = "Envelopes";
            const string Force = "Force ";
            const string Fractals = "Fractals";
            const string Gator = "Gator";
            const string Ichimoku = "Ichimoku";
            const string BWMF = "BWMF";
            const string Momentum = "Momentum";
            const string MFI = "MFI";
            const string MA = "MA";
            const string OSMA = "OSMA";
            const string MACD = "MACD";
            const string OBV = "OBV";
            const string SAR = "SAR";
            const string RSI = "RSI";
            const string RVI = "RVI";
            const string StdDev = "StdDev";
            const string Stochastic = "Stochastic";
            const string WPR = "WPR";

            string gpu = InputSheet.Cells[3, 2].Value2;             // GPU - Enable / Disable
            string isTrain = InputSheet.Cells[4, 2].Value2;         // Train - Yes / No
            string Re_Train = InputSheet.Cells[5, 2].Value2;        // Re-Train - Yes / No
            string En_Indi = InputSheet.Cells[19, 2].Value2;        // Enable Indicator - Enable / Disable

            string FileName = InputSheet.Cells[8, 2].Value2;
            string Training_data = InputSheet.Cells[6, 2].Value2;          // Training Data - O, H, L, C
            string Architecture = InputSheet.Cells[14, 2].Value2;          // Training Model - LSTM / GRU
            string Optimizer = InputSheet.Cells[15, 2].Value2;             // Optimizer - MSQ / RSQ / CORL
            string Ch_Indi = InputSheet.Cells[20, 2].Value2;               // Choose a Indicator( applied when Indicator is Enabled)
            string Reg_cls = InputSheet.Cells[21, 2].Value2;               // Regression / classification

            double learningrate = (double)InputSheet.Cells[9, 2].Value2;   // Learning Rated
            double testingPart = (double)InputSheet.Cells[12, 2].Value2;   // Testing Part (in %)
            double testingWeight = (double)InputSheet.Cells[13, 2].Value2;  // Testing Weight (in %)
            double momentum = (double)InputSheet.Cells[16, 2].Value2;      // Momentum
            int Epochs = (int)InputSheet.Cells[10, 2].Value2;              // Epochs
            int Scale = (int)InputSheet.Cells[11, 2].Value2;               // Scale             
            int Maxbars = (int)InputSheet.Cells[18, 2].Value2;             // Maxbars
            int Minbars = (int)InputSheet.Cells[17, 2].Value2;             // Minbars

            object[,] price = DataRange3.Value2;
            OutRange3.Value2 = price;

            object[,] date = DataSheet.Range[DataSheet.Cells[2, 1], DataSheet.Cells[(int)size + 1, 1]].Value2;
            object[,] time = DataSheet.Range[DataSheet.Cells[2, 2], DataSheet.Cells[(int)size + 1, 2]].Value2;

            double[] trainingData = new double[(int)size];
            for (int i = 1; i < size + 1; i++)
            {
                trainingData[i - 1] = Double.Parse(price[i, 1].ToString());
            }

            long[] dateArray = new long[(int)size];
            for (int i = 1; i < size + 1; i++)
            {
                dateArray[i - 1] = (Int64)(DateTime.Parse(string.Join(" ", date[i, 1].ToString(), TimeSpan.FromDays(Double.Parse(time[i, 1].ToString
                    ()))).ToString()).Subtract(new DateTime(1970, 1, 1, 5, 30, 0)).TotalSeconds);
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("train;");

            foreach (double t_data_ele in trainingData)
            {
                sb.Append(t_data_ele);
                sb.Append(',');
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append(';');

            foreach (long t_date_ele in dateArray)
            {
                sb.Append(t_date_ele);
                sb.Append(',');
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append(';');

            sb.Append(FileName + ';');
            sb.Append(Epochs.ToString() + ';');
            sb.Append(learningrate.ToString() + ';');
            sb.Append(momentum.ToString() + ';');
            sb.Append(Scale.ToString() + ';');
            sb.Append((Optimizer.Equals("RMSProp") ? 1 : 0).ToString() + ';');
            sb.Append(testingPart.ToString() + ';');
            sb.Append(testingWeight.ToString());

            MultiCharts multiCharts = new MultiCharts
            {
                action = "train",
                gpu = gpu.Equals("Yes") ? true : false,
                data = trainingData,
                date = dateArray,
                fileName = FileName,
                epochs = Epochs,
                learningRate = learningrate,
                momentum = momentum,
                scale = Scale,
                optimizer = Optimizer.Equals("RMSProp") ? 1 : 0,
                testingPart = testingPart,
                testingWeight = testingWeight
            };

            Directory.SetCurrentDirectory("C:\\MultiCharts");

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(multiCharts);
            File.WriteAllText(Path.Combine(Directory.GetCurrentDirectory(), FileName + ".json"), json);

            Process.Start(Path.Combine(Directory.GetCurrentDirectory(), "MultiChartsClientCS.exe"), sb.ToString());

            if (isTrain == "Yes")
            {
                if (Architecture == "LSTM")
                {
                    OutRange1.Value2 = date;
                    OutRange2.Value2 = time;
                    ErrorRange1.Value2 = date;
                    ErrorRange2.Value2 = time;

                    switch (Training_data)
                    {
                        case open:                              // Case for "OPEN" Training Data

                            break;

                        case high:                              // Case for "HIGH" Training Data 
                            price = DataRange4.Value2;
                            OutRange4.Value2 = price;
                            break;

                        case low:                               // Case for "LOW" Training Data
                            price = DataRange5.Value2;
                            OutRange5.Value2 = price;
                            break;

                        case close:                             // Case for "CLOSE" Training Data
                            price = DataRange6.Value2;
                            OutRange6.Value2 = price;
                            break;
                    }
                }
                
                /******  Error Analysis  ******/



                for (int j = 3; j < 7; j++)
                {
                    for (int i = 2; i <= (int)size + 1; i++)
                    {
                        if (OutputSheet.Cells[i, j].Value2 == null)
                            break;
                        ErrorSheet.Cells[i, j].Value2 = DataSheet.Cells[i, j].Value2 - OutputSheet.Cells[i, j].Value2;
                    }
                }

                /*****  Indicator  *****/

                if (En_Indi == "Enable")
                {
                    if (Ch_Indi != OutputSheet.Cells[1, 7].value2)
                    {
                        object[,] value;
                        switch (Ch_Indi)
                        {
                            case AC:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 3], IndicatorSheet.Cells[(int)size + 1, 3]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case AD:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 4], IndicatorSheet.Cells[(int)size + 1, 4]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case ADX:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 5], IndicatorSheet.Cells[(int)size + 1, 5]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Alligator:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 6], IndicatorSheet.Cells[(int)size + 1, 6]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case AO:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 7], IndicatorSheet.Cells[(int)size + 1, 7]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case ATR:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 8], IndicatorSheet.Cells[(int)size + 1, 8]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case BearsPower:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 9], IndicatorSheet.Cells[(int)size + 1, 9]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Bands:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 10], IndicatorSheet.Cells[(int)size + 1, 10]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case BullsPower:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 11], IndicatorSheet.Cells[(int)size + 1, 11]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case CCI:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 12], IndicatorSheet.Cells[(int)size + 1, 12]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Custom:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 13], IndicatorSheet.Cells[(int)size + 1, 13]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case DeMarker:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 14], IndicatorSheet.Cells[(int)size + 1, 14]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Envelopes:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 15], IndicatorSheet.Cells[(int)size + 1, 15]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Force:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 16], IndicatorSheet.Cells[(int)size + 1, 16]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Fractals:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 17], IndicatorSheet.Cells[(int)size + 1, 17]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Gator:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 18], IndicatorSheet.Cells[(int)size + 1, 18]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Ichimoku:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 19], IndicatorSheet.Cells[(int)size + 1, 19]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case BWMF:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 20], IndicatorSheet.Cells[(int)size + 1, 20]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Momentum:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 21], IndicatorSheet.Cells[(int)size + 1, 21]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case MFI:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 22], IndicatorSheet.Cells[(int)size + 1, 22]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case MA:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 23], IndicatorSheet.Cells[(int)size + 1, 23]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case OSMA:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 24], IndicatorSheet.Cells[(int)size + 1, 24]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case MACD:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 25], IndicatorSheet.Cells[(int)size + 1, 25]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case OBV:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 26], IndicatorSheet.Cells[(int)size + 1, 26]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case SAR:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 27], IndicatorSheet.Cells[(int)size + 1, 27]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case RSI:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 28], IndicatorSheet.Cells[(int)size + 1, 28]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case RVI:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 29], IndicatorSheet.Cells[(int)size + 1, 29]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case StdDev:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 30], IndicatorSheet.Cells[(int)size + 1, 30]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case Stochastic:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 31], IndicatorSheet.Cells[(int)size + 1, 31]].Value2;
                                OutRange7.Value2 = value;
                                break;

                            case WPR:
                                value = IndicatorSheet.Range[IndicatorSheet.Cells[1, 32], IndicatorSheet.Cells[(int)size + 1, 32]].Value2;
                                OutRange7.Value2 = value;
                                break;

                        }
                    }
                }
            }
        }
        
        /************  EVALUATE BUTTON  ************/

        private void Evaluate_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook actbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            Excel.Worksheet InputSheet = actbook.Sheets[1];

            string FileName = InputSheet.Cells[8, 2].Value2;
            double testingWeight = (double)InputSheet.Cells[13, 2].Value2;  // Testing Weight (in %)

            StringBuilder sb = new StringBuilder();
            sb.Append("eval;");

            sb.Append(testingWeight.ToString() + ';');
            sb.Append(FileName);

            Directory.SetCurrentDirectory("C:\\MultiCharts");
            Process.Start(Path.Combine(Directory.GetCurrentDirectory(), "MultiChartsClientCS.exe"), sb.ToString());
        }

        /************  FORECAST BUTTON  ************/

        private void Forecast_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook actbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            Excel.Worksheet InputSheet = actbook.Sheets[1];

            Excel.Worksheet DataSheet = actbook.Sheets[2];

            double size = InputSheet.Cells[7, 2].Value2;
           
            int ticks = (int)InputSheet.Cells[22, 2].Value2;               // Number of Bars to forecast
            string FileName = InputSheet.Cells[8, 2].Value2;

            object[,] date = DataSheet.Range[DataSheet.Cells[2, 1], DataSheet.Cells[(int)size + 1, 1]].Value2;
            object[,] time = DataSheet.Range[DataSheet.Cells[2, 2], DataSheet.Cells[(int)size + 1, 2]].Value2;

            long lastDateTimeMinusOne = (Int64)(DateTime.Parse(string.Join(" ", date[(int)size - 2, 1].ToString(), TimeSpan.FromDays(Double.Parse(time[(int)size - 2, 1].ToString())))).Subtract(new DateTime(1970, 1, 1, 5, 30, 0)).TotalSeconds);

            long lastDateTime = (Int64)(DateTime.Parse(string.Join(" ", date[(int)size - 1, 1].ToString(), TimeSpan.FromDays(Double.Parse(time[(int)size - 1, 1].ToString())))).Subtract(new DateTime(1970, 1, 1, 5, 30, 0)).TotalSeconds);

            StringBuilder sb = new StringBuilder();
            sb.Append("forecast;");            
            sb.Append(ticks.ToString() + ';');
            sb.Append(lastDateTime.ToString() + ';');
            sb.Append((lastDateTime - lastDateTimeMinusOne).ToString() + ';');
            sb.Append(FileName);

            Directory.SetCurrentDirectory("C:\\MultiCharts");
            Process.Start(Path.Combine(Directory.GetCurrentDirectory(), "MultiChartsClientCS.exe"), sb.ToString());

            Prediction_Label.Label = "Predicted Values : ";
            Display_Forecast.Visible = true;
            Prediction_Label.Visible = true;
        }

        /**************  DELETE BUTTON  ***************/

        private void Delete_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet InputSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[1];

            Excel.Worksheet OutputSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[4];

            Excel.Worksheet ErrorSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[6];

            double size = InputSheet.Cells[7, 2].Value2;

            Excel.Range OutRange = OutputSheet.Range[OutputSheet.Cells[2, 1], OutputSheet.Cells[(int)size + 1, 7]];

            Excel.Range ErrorRange = ErrorSheet.Range[ErrorSheet.Cells[2, 1], ErrorSheet.Cells[(int)size + 1, 7]];

            OutputSheet.Cells[1, 7].ClearContents();
            OutRange.Clear();
            ErrorRange.Clear();

        }

        /**************  RETRIEVE BUTTON  ***************/

        private void Retrieve_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook actbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            Excel.Worksheet DataSheet = actbook.Sheets[2];

            int lastUsedRow = DataSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            while (DataSheet.Cells[lastUsedRow, 1].Value2 == null)
                lastUsedRow--;

            String fileData = System.IO.File.ReadAllText(@"");

            String[] lines = fileData.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            int totalRowsInFile = lines.Length;

            int index = lastUsedRow;

            while (index < totalRowsInFile)
            {
                String[] values = lines[index].Split(',');

                for (int i = 0; i < 6; i++)
                {
                    DataSheet.Cells[index + 1, i+1] = values[i];
                }

                index++;
            }
        }
    }
}