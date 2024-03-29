﻿using System;
using System.IO;
using System.Reflection;

#region #usings_Sh
using DevExpress.Spreadsheet;
#endregion #usings_Sh

namespace SpreadsheetControl_WPF_API_Part02
{
    public static class ShapeActions {
        #region Actions
        public static Action<IWorkbook> InsertShapeAction = InsertShapeValue;
        public static Action<IWorkbook> InsertShapeFromUriAction = InsertShapeFromUriValue;
        public static Action<IWorkbook> ModifyShapeAction = ModifyShapeValue;
        #endregion

        static void InsertShapeValue(IWorkbook workbook)
        {
            #region #insertshape
            Stream imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SpreadsheetControl_WPF_API_Part02.Pictures.chart.png");
            SpreadsheetImageSource imageSource = SpreadsheetImageSource.FromStream(imageStream);
            workbook.BeginUpdate();
            // Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert a picture from a file so that its top left corner is in the specified cell.
                // By default the picture is named Picture1.. PictureNN.
                worksheet.Pictures.AddPicture("Pictures\\chart.png", worksheet.Cells["D5"]);
                // Insert a picture to fit in the specified range.
                worksheet.Pictures.AddPicture("Pictures\\chart.png", worksheet.Range["B2"]);
                // Insert a picture from the SpreadsheetImageSource at 120 mm from the left, 80 mm from the top, 
                // and resize it to a width of 70 mm and a height of 20 mm, locking the aspect ratio.
                worksheet.Pictures.AddPicture(imageSource, 120, 80, 70, 20, true);
                // Insert the picture to be removed.
                worksheet.Pictures.AddPicture("Pictures\\chart.png", 0, 0);
                // Remove the last inserted picture.
                // Find the shape by its name. The method returns a collection of shapes with the same name.
                Shape picShape = worksheet.Shapes.GetShapesByName("Picture 4")[0];
                picShape.Delete();
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #insertshape
        }

        static void InsertShapeFromUriValue(IWorkbook workbook)
        {
            #region #insertshapefromuri
            string imageUri = "https://docs.devexpress.com/WPF/images/spreadsheet-main-page.png";
            // Create an image from Uri.
            SpreadsheetImageSource imageSource = SpreadsheetImageSource.FromUri(imageUri, workbook);
            // Set the measurement unit to point.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;
                        
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert a picture from the SpreadsheetImageSource at 100 pt from the left, 40 pt from the top, 
                // and resize it to a width of 200 pt and a height of 180 pt.
                worksheet.Pictures.AddPicture(imageSource, 100, 40, 400, 300 );
            }
            finally
            {
                workbook.EndUpdate();
            }

            #endregion #insertshapefromuri
            
        }

        static void ModifyShapeValue(IWorkbook workbook)
        {
            #region #modifyshape
            // Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert pictures from the file.
                Picture pic = worksheet.Pictures.AddPicture("Pictures\\chart.png", worksheet.Cells["A1"]);
                worksheet.Pictures.AddPicture("Pictures\\candles.png", worksheet.Cells["D5"]);
                // Specify picture name and draw a border.
                pic.Name = "Candles";
                pic.AlternativeText = "Candles snapshot";
                pic.BorderWidth = 1;
                pic.BorderColor = DevExpress.Utils.DXColor.Black;
                // Move a picture.
                pic.Move(20, 30);
                // Change picture behavior so it will move and size with underlying cells. 
                pic.Placement = Placement.MoveAndSize;
                worksheet.Rows[5].Height += 10;
                worksheet.Columns["D"].Width += 10;
                // Specify rotation angle.
                pic.Rotation = 30;
                // Increase rotation angle.
                pic.IncrementRotation(-30);
                // Add a hyperlink.
                pic.InsertHyperlink("http://www.devexpress.com/", true);
                worksheet.Shapes[1].InsertHyperlink("http://www.devexpress.com/Products/NET/Controls/WPF/", true);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #modifyshape
        }



    }
}
