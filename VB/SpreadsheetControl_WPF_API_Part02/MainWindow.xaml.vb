Imports DevExpress.Spreadsheet
Imports DevExpress.Xpf.NavBar
Imports System
Imports System.Net
Imports System.Windows
Imports System.Windows.Input

Namespace SpreadsheetControl_WPF_API_Part02
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits Window

        Private workbook As IWorkbook

        Public Sub New()
            System.Net.ServicePointManager.SecurityProtocol = System.Net.ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12
            InitializeComponent()
            ' Access a workbook.
            workbook = spreadsheetControl1.Document

            DataContext = Groups.InitData()

        End Sub

        Private Sub NavigationPaneView_MouseDoubleClick(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            Dim item As NavBarItem = DirectCast(sender, NavBarViewBase).GetNavBarItem(e)
            If item IsNot Nothing Then
                Dim example As SpreadsheetExample = TryCast(item.Content, SpreadsheetExample)
                If example IsNot Nothing Then
                    workbook.Options.Culture = System.Globalization.CultureInfo.CurrentCulture
                    LoadDocumentFromFile()
                    example.Action()(workbook)
                    SaveDocumentToFile()
                End If
            End If
        End Sub
        ' ------------------- Load and Save a Document -------------------
        Private Sub LoadDocumentFromFile()
            '            #Region "#LoadDocumentFromFile"
            ' Load a workbook from a file.
            workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
            '            #End Region ' #LoadDocumentFromFile
        End Sub

        Private Sub SaveDocumentToFile()
            '            #Region "#SaveDocumentToFile"
            ' Save the modified document to a file.
            workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
'            #End Region ' #SaveDocumentToFile
        End Sub

    End Class
End Namespace
