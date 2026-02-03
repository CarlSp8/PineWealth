using System.CodeDom.Compiler;
using System.Windows;
using Microsoft.Win32;
using WealthLab.Backtest;
using WealthLab.Core;

namespace PineWealth
{
    public partial class MainWindow : Window
    {
        //constructor
        public MainWindow()
        {
            InitializeComponent();
        }

        //window loaded
        private void onLoad(object sender, RoutedEventArgs e)
        {
            //create host
            WLHost.Instance = new PSTHost();

            //load default PineScript into editor
            txtPineScript.Text = Properties.Resources.PineScript;
        }

        //perform transdlation
        private void btnTranslateClick(object sender, RoutedEventArgs e)
        {
            txtStatus.Clear();
            PineScriptTranslator pst = new PineScriptTranslator();
            try
            {
                txtCSharp.Text = pst.Translate(txtPineScript.Text);
                txtStatus.AppendText("Translation succeeded." + Environment.NewLine);

                //attempt to compile the strategy code
                Strategy s = new Strategy(StrategyType.Code);
                s.StrategyData = txtCSharp.Text;
                UserStrategyBase usb = s.CreateNewInstance();
                if (usb != null)
                    txtStatus.AppendText("Strategy code compiled OK." + Environment.NewLine);

                //log errors
                if (s.CompilerErrors.Count > 0)
                {
                    foreach (CompilerError ce in s.CompilerErrors)
                        txtStatus.AppendText(ce.ErrorText + Environment.NewLine);
                }
            }
            catch(Exception ex)
            {
                txtStatus.AppendText(ex.Message + Environment.NewLine);
                txtStatus.AppendText(ex.StackTrace);
            }
        }

        //export to Excel
        private void btnExportExcelClick(object sender, RoutedEventArgs e)
        {
            //check if there's C# code to export
            if (string.IsNullOrWhiteSpace(txtCSharp.Text))
            {
                MessageBox.Show("Please translate the Pine Script first before exporting to Excel.", 
                    "No Translation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            //show save file dialog
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = "PineWealth_Translation"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    ExcelExporter exporter = new ExcelExporter();
                    exporter.Export(txtPineScript.Text, txtCSharp.Text, saveFileDialog.FileName);
                    txtStatus.AppendText($"Excel file saved to: {saveFileDialog.FileName}" + Environment.NewLine);
                    MessageBox.Show($"Excel file exported successfully!\n\nSaved to: {saveFileDialog.FileName}", 
                        "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    txtStatus.AppendText("Error exporting to Excel: " + ex.Message + Environment.NewLine);
                    MessageBox.Show($"Error exporting to Excel:\n{ex.Message}", 
                        "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}