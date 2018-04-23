﻿Imports Microsoft.VisualBasic
Imports DevExpress.Spreadsheet
Imports DevExpress.Xpf.NavBar
Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Navigation
Imports System.Windows.Shapes
Imports DevExpress.Xpf.Spreadsheet

Namespace SpreadsheetControl_WPF_API_Part02
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits Window
        Private workbook As IWorkbook

        Public Sub New()
            InitializeComponent()
            ' Access a workbook.
            workbook = spreadsheetControl1.Document

            DataContext = Groups.InitData()

        End Sub

        Private Sub NavigationPaneView_MouseDoubleClick(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            Dim item As NavBarItem = (CType(sender, NavBarViewBase)).GetNavBarItem(e)
            If item IsNot Nothing Then
                Dim example As SpreadsheetExample = TryCast(item.Content, SpreadsheetExample)
                If example IsNot Nothing Then
                    Dim action As Action(Of IWorkbook) = example.Action
                    LoadDocumentFromFile()
                    action(workbook)
                    SaveDocumentToFile()
                End If
            End If
        End Sub
        ' ------------------- Load and Save a Document -------------------
        Private Sub LoadDocumentFromFile()
            '			#Region "#LoadDocumentFromFile"
            ' Load a workbook from a file.
            workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
            '			#End Region ' #LoadDocumentFromFile
        End Sub

        Private Sub SaveDocumentToFile()
            '			#Region "#SaveDocumentToFile"
            ' Save the modified document to a file.
            workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
            '			#End Region ' #SaveDocumentToFile
        End Sub

    End Class
End Namespace
