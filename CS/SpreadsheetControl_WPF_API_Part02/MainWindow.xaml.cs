using DevExpress.Spreadsheet;
using DevExpress.Xpf.NavBar;
using System;
using System.Windows;
using System.Windows.Input;

namespace SpreadsheetControl_WPF_API_Part02
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : DevExpress.Xpf.Core.ThemedWindow
    {
        IWorkbook workbook;

        public MainWindow()
        {
            InitializeComponent();
            // Access a workbook.
            workbook = spreadsheetControl1.Document;

            DataContext = Groups.InitData();

        }

        private void NavigationPaneView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            NavBarItem item = ((NavBarViewBase)sender).GetNavBarItem(e);
            if (item != null)
            {
                SpreadsheetExample example = item.Content as SpreadsheetExample;
                if (example != null)
                {
                    workbook.Options.Culture = System.Globalization.CultureInfo.CurrentCulture; 
                    LoadDocumentFromFile();
                    example.Action(workbook);
                    SaveDocumentToFile();
                }
            }
        }
        // ------------------- Load and Save a Document -------------------
        private void LoadDocumentFromFile()
        {
            #region #LoadDocumentFromFile
            // Load a workbook from a file.
            workbook.LoadDocument("Documents\\Document.xlsx", DocumentFormat.OpenXml);
            #endregion #LoadDocumentFromFile
        }

        private void SaveDocumentToFile()
        {
            #region #SaveDocumentToFile
            // Save the modified document to a file.
            workbook.SaveDocument("Documents\\SavedDocument.xlsx", DocumentFormat.OpenXml);
            #endregion #SaveDocumentToFile
        }

    }
}
