namespace ExcelAddIn5
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Trade = this.Factory.CreateRibbonGroup();
            this.Train = this.Factory.CreateRibbonButton();
            this.Evaluate = this.Factory.CreateRibbonButton();
            this.Forecast = this.Factory.CreateRibbonButton();
            this.Manage_Data = this.Factory.CreateRibbonGroup();
            this.Retrieve_Data = this.Factory.CreateRibbonButton();
            this.Clear_Output = this.Factory.CreateRibbonButton();
            this.Display_Forecast = this.Factory.CreateRibbonGroup();
            this.Prediction_Label = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.Trade.SuspendLayout();
            this.Manage_Data.SuspendLayout();
            this.Display_Forecast.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.Trade);
            this.tab1.Groups.Add(this.Manage_Data);
            this.tab1.Groups.Add(this.Display_Forecast);
            this.tab1.Label = "HPCS";
            this.tab1.Name = "tab1";
            // 
            // Trade
            // 
            this.Trade.Items.Add(this.Train);
            this.Trade.Items.Add(this.Evaluate);
            this.Trade.Items.Add(this.Forecast);
            this.Trade.Label = "Trade";
            this.Trade.Name = "Trade";
            // 
            // Train
            // 
            this.Train.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Train.Image = ((System.Drawing.Image)(resources.GetObject("Train.Image")));
            this.Train.Label = "Train";
            this.Train.Name = "Train";
            this.Train.ShowImage = true;
            this.Train.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Train_Click);
            // 
            // Evaluate
            // 
            this.Evaluate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Evaluate.Image = ((System.Drawing.Image)(resources.GetObject("Evaluate.Image")));
            this.Evaluate.Label = "Evaluate";
            this.Evaluate.Name = "Evaluate";
            this.Evaluate.ShowImage = true;
            this.Evaluate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Evaluate_Click);
            // 
            // Forecast
            // 
            this.Forecast.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Forecast.Image = ((System.Drawing.Image)(resources.GetObject("Forecast.Image")));
            this.Forecast.Label = "Forecast";
            this.Forecast.Name = "Forecast";
            this.Forecast.ShowImage = true;
            this.Forecast.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Forecast_Click);
            // 
            // Manage_Data
            // 
            this.Manage_Data.Items.Add(this.Retrieve_Data);
            this.Manage_Data.Items.Add(this.Clear_Output);
            this.Manage_Data.Label = "Manage Data";
            this.Manage_Data.Name = "Manage_Data";
            // 
            // Retrieve_Data
            // 
            this.Retrieve_Data.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Retrieve_Data.Image = ((System.Drawing.Image)(resources.GetObject("Retrieve_Data.Image")));
            this.Retrieve_Data.Label = "Retrieve Data";
            this.Retrieve_Data.Name = "Retrieve_Data";
            this.Retrieve_Data.ShowImage = true;
            this.Retrieve_Data.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Retrieve_Click);
            // 
            // Clear_Output
            // 
            this.Clear_Output.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Clear_Output.Image = ((System.Drawing.Image)(resources.GetObject("Clear_Output.Image")));
            this.Clear_Output.Label = "Clear Output";
            this.Clear_Output.Name = "Clear_Output";
            this.Clear_Output.ShowImage = true;
            this.Clear_Output.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Delete_Click);
            // 
            // Display_Forecast
            // 
            this.Display_Forecast.Items.Add(this.Prediction_Label);
            this.Display_Forecast.Label = "Predicted Values";
            this.Display_Forecast.Name = "Display_Forecast";
            this.Display_Forecast.Visible = false;
            // 
            // Prediction_Label
            // 
            this.Prediction_Label.Label = "Predictions";
            this.Prediction_Label.Name = "Prediction_Label";
            this.Prediction_Label.Visible = false;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Trade.ResumeLayout(false);
            this.Trade.PerformLayout();
            this.Manage_Data.ResumeLayout(false);
            this.Manage_Data.PerformLayout();
            this.Display_Forecast.ResumeLayout(false);
            this.Display_Forecast.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Trade;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Forecast;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Manage_Data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Clear_Output;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Train;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Retrieve_Data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Evaluate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Display_Forecast;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel Prediction_Label;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
