using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace ExportComponent
{
    public class ExportExcel
    {
        public void ExportToExcel(DataSet ds, List<string> SheetName)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel._Worksheet)workbook.ActiveSheet;
            app.Visible = true;
            app.DisplayAlerts = false;
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                System.Data.DataTable dt = ds.Tables[SheetName[i]];
                worksheet = (Microsoft.Office.Interop.Excel._Worksheet)app.Worksheets.Add();
                worksheet.Name = SheetName[i];
                #region header
                Range b11 = worksheet.Range["A1", "A1"];
                b11.Value = "TÊN VẬT TƯ";
                b11.Font.Bold = true;

                Range c11 = worksheet.Range["B1", "B1"];
                c11.Value = "QUI CÁCH";
                c11.Font.Bold = true;

                Range d11 = worksheet.Range["C1", "C1"];
                d11.Value = "ĐVT";
                d11.Font.Bold = true;

                Range e11 = worksheet.Range["D1", "D1"];
                e11.Value = "SỐ LƯỢNG";
                e11.Font.Bold = true;

                Range f11 = worksheet.Range["E1", "E1"];
                f11.Value = "KHỐI LƯỢNG 1/CK";
                f11.Font.Bold = true;

                Range g11 = worksheet.Range["F1", "F1"];
                g11.Value = "TỔNG KHỐI LƯỢNG";
                g11.Font.Bold = true;
                #endregion
                object[,] array2 = new object[dt.Rows.Count, dt.Columns.Count];
                for (int ii = 0; ii < dt.Rows.Count; ii++)
                {
                    DataRow dataRow2 = dt.Rows[ii];
                    for (int jj = 0; jj < dt.Columns.Count; jj++)
                    {
                        array2[ii, jj] = dataRow2[jj];
                    }
                }
                #region Xác định vùng đổ dữ liệu
                int rowStart2 = 2;
                int columnStart2 = 1;
                int rowEnd2 = rowStart2 + dt.Rows.Count - 1;
                int columnEnd2 = dt.Columns.Count;
                #endregion
                #region format

                // Ô bắt đầu điền dữ liệu
                Microsoft.Office.Interop.Excel.Range cc3 =
                    (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[rowStart2, columnStart2];
                // Ô kết thúc điền dữ liệu
                Microsoft.Office.Interop.Excel.Range cc4 =
                    (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[rowEnd2, columnEnd2];
                // Lấy về vùng điền dữ liệu
                Microsoft.Office.Interop.Excel.Range range2 = worksheet.get_Range(cc3, cc4);
                //Điền dữ liệu vào vùng đã thiết lập
                range2.Value = array2;

                // Ô bắt đầu fm
                Microsoft.Office.Interop.Excel.Range cc2 =
                    (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[rowStart2, columnStart2];
                // Ô kết thúc fm
                Microsoft.Office.Interop.Excel.Range cclast =
                    (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[rowEnd2, columnStart2];
                // Căn giữa cột STT
                worksheet.get_Range(b11, cc4).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                worksheet.get_Range(cc2, cclast).HorizontalAlignment =
    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                workbook.Windows[1].WindowState = XlWindowState.xlMaximized;
                //Định dạng kẻ viền
                Range range55 = worksheet.get_Range(b11, cc4);
                range55.Borders.LineStyle = XlLineStyle.xlContinuous;
                range55.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic);
                worksheet.Columns.AutoFit();
                #endregion
            }
            #region Save file
            //string textPath = "e:\\output.xls";
            //if (File.Exists(textPath))
            //{
            //    File.Delete(textPath);
            //}

            //workbook.SaveAs(textPath, Type.Missing, Type.Missing, Type.Missing,
            // Type.Missing,
            // Type.Missing,
            // Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
            // Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //app.Quit();
            #endregion


        }
    }
}
