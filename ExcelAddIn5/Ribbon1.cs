using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelAddIn5
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
                                                                /*******    DLL Import   ********/
                                                               

        private const string dllAddress = "C:\\Users\\HPCS\\Documents\\Visual Studio 2015\\Projects\\MultiChartsProject\\MultiChartsDLL\\x64\\Release\\MultiChartsDLL.dll";

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "??1MultiCharts@@QEAA@XZ")]
        public static extern void DisposeMultiCharts();

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?InitTrainingData@MultiCharts@@QEAAXH@Z")]
        public static extern void InitTrainingData( int size);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?SetTrainingData@MultiCharts@@QEAAXPEAN@Z")]
        public static extern void SetTrainingData(double[] trainingData);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?InitDateArray@MultiCharts@@QEAAXH@Z")]
        public static extern double InitDateArray( int size);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?InitFileName@MultiCharts@@QEAAXH@Z")]
        public static extern double InitFileName(int size);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?InitVolumeArray@MultiCharts@@QEAAXH@Z")]
        public static extern double InitVolumeArray(int size);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?SetDateArray@MultiCharts@@QEAAXPEAD@Z")]
        public static extern void SetDateArray( char[] dateArray);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?SetEpochs@MultiCharts@@QEAAXH@Z")]
        public static extern void SetEpochs(int epochs);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?SetFileName@MultiCharts@@QEAAXPEAD@Z")]
        public static extern void SetFileName(char filename);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?SetLearningRate@MultiCharts@@QEAAXN@Z")]
        public static extern void SetlearningRate(double learningRate);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?SetMomentum@MultiCharts@@QEAAXH@Z")]
        public static extern void SetMomentum(int momentum);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?SetOptimizer@MultiCharts@@QEAAXH@Z")]
        public static extern void SetOptimizer(int optimizer);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "??SetScale@MultiCharts@@QEAAXH@Z")]
        public static extern void SetScale(int scale );

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?SetVolumeArray@MultiCharts@@QEAAXPEAJ@Z")]
        public static extern void SetVolumeArray(long volume);

        [DllImport(dllAddress, CallingConvention = CallingConvention.Cdecl, EntryPoint = "?TrainModel@MultiCharts@@QEAANXZ")]
        public static extern double TrainModel();

        

                                                                /************  FORECAST BUTTON  ************/

           
        private void Forecast_Click(object sender, RibbonControlEventArgs e)
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
            int x = 0;

            string gpu = InputSheet.Cells[3, 2].Value2;             // GPU - Enable / Disable
            string Train = InputSheet.Cells[4, 2].Value2;           // Train - Yes / No
            string Re_Train = InputSheet.Cells[5, 2].Value2;        // Re-Train - Yes / No
            string En_Indi = InputSheet.Cells[18, 2].Value2;        // Enable Indicator - Enable / Disable

            string Training_data = InputSheet.Cells[6, 2].Value2;   // Training Data - O, H, L, C
            string Mode = InputSheet.Cells[13, 2].Value2;           // Training Mode - LSTM / GRNN / BP
            string Model = InputSheet.Cells[14, 2].Value2;          // Training Model - MSQ / RSQ / CORL
            string Ch_Indi = InputSheet.Cells[19, 2].Value2;        // Choose a Indicator( applied when Indicator is Enabled)
            string Reg_cls = InputSheet.Cells[20, 2].Value2;        // Regression / classification
            
            double learningrate = (double)InputSheet.Cells[9, 2].Value2;     // Learning Rate
            int Epochs = (int)InputSheet.Cells[10, 2].Value2;            // Epochs
            int Scale = (int)InputSheet.Cells[11, 2].Value2;             // Scale                                  // issue - Output don't comes up (D,T,O, H, L,C,I)
            int Optimizer = (int)InputSheet.Cells[12, 2].Value2;         // Optimizer
            int momentum = (int)InputSheet.Cells[15, 2].Value2;          // Momentum
            int Maxbars = (int)InputSheet.Cells[16, 2].Value2;           // Maxbars
            int Minbars = (int)InputSheet.Cells[17, 2].Value2;           // Minbars
            
          
            object[,] price = new object[(int)size, 1];
            object[,] date = new object[(int)size, 1];
            object[,] time = new object[(int)size, 1];
            object[,] value = new object[(int)size, 1];

            if (Train == "Yes")
            {
                date = DataSheet.Range[DataSheet.Cells[2, 1], DataSheet.Cells[(int)size + 1, 1]].Value2;
                time = DataSheet.Range[DataSheet.Cells[2, 2], DataSheet.Cells[(int)size + 1, 2]].Value2;

            Start:
                if (Mode == "LSTM")
                {
                    OutRange1.Value2 = date;
                    OutRange2.Value2 = time;
                    ErrorRange1.Value2 = date;
                    ErrorRange2.Value2 = time;
                   
                    switch (Training_data)
                    {
                        case open:                              // Case for "OPEN" Training Data
                            price = DataRange3.Value2;
                            OutRange3.Value2 = price;
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

                if (Mode == "GRNN")
                {
                    OutRange1.Value2 = date;
                    OutRange2.Value2 = time;
                    ErrorRange1.Value2 = date;
                    ErrorRange2.Value2 = time;

                    switch (Training_data)
                    {
                        case open:                              // Case for "OPEN" Training Data
                            price = DataRange3.Value2;
                            OutRange3.Value2 = price;
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
               
                if (Mode == "BP")
                {
                    OutRange1.Value2 = date;
                    OutRange2.Value2 = time;
                    ErrorRange1.Value2 = date;
                    ErrorRange2.Value2 = time;

                    switch (Training_data)
                    {
                        case open:                              // Case for "OPEN" Training Data
                            price = DataRange3.Value2;
                            OutRange3.Value2 = price;
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
                                                                    /*****  DLL Calls  *****/
                
                                
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

                                                                         /*****  GPU  *****/
                if (gpu == "Enabled")
                {

                }
                                                                        /*****  Re-Train  *****/
                if (x == 5)
                {
                    x = 0;
                    goto end;
                }
                                                                      
                if (Re_Train == "Yes")
                {
                    x = 5;
                    goto Start;
                }

            end:
                /*****  Indicator  *****/

                if (En_Indi == "Enable")
                {
                    if (Ch_Indi != OutputSheet.Cells[1, 7].value2)
                    {
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

                /*********   Regression / Classification  *********/
                if (Reg_cls == "Regression")
                {

                }

                if (Reg_cls == "Classification")
                {
                    if((OutputSheet.Cells[(int)size +1, 4 ]).Value2 > (OutputSheet.Cells[(int)size + 1, 4]).Value2)           // Check Close price
                    { }
                    else{ }
                }
                                                              /*********  Training Model  *************/
                if (Model == "MSQ")
                {

                }

                if (Model == "RSQ")
                {

                }

                if (Model == "CORL")
                {

                }
                InitTrainingData((int)size);
                InitDateArray((int)size);                            // issue -indicator values don't comes up on output sheet(if uncommented)
                InitFileName((int)size);
                InitVolumeArray((int)size);

                
                SetlearningRate(learningrate);
                SetEpochs(Epochs);
                SetScale(Scale);                                      
                SetOptimizer(Optimizer);
                SetMomentum(momentum);
               



            }
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
    }
}