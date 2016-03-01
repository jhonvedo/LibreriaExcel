using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelLibrary.Writer
{
    public class ExcelWriter
    {
        #region PROPIEDADES

        private Application xlApp;
        private Workbook xlBook;
        private Worksheet xlSheet;
        private string ruta;

        #endregion PROPIEDADES

        #region ATRIBUTOS

        public Application App
        {
            get { return xlApp; }
            set { xlApp = value; }
        }

        public Workbook Book
        {
            get { return xlBook; }
            set { xlBook = value; }
        }

        public Worksheet Sheet
        {
            get { return xlSheet; }
            set { xlSheet = value; }
        }

        public string Ruta
        {
            get { return ruta; }
            set { ruta = value; }
        }

        #endregion ATRIBUTOS

        #region CONSTRUCTOR

        /// <summary>
        /// Constuctor por defector,
        /// </summary>
        public ExcelWriter()
        {
            xlApp = new Application();
            object misValue = System.Reflection.Missing.Value;
            xlBook = xlApp.Workbooks.Add(misValue);
            xlSheet = (Worksheet)xlBook.Worksheets.get_Item(1);
        }

        #endregion CONSTRUCTOR

        #region MÉTODOS PÚBLICOS

        //public void CloseBook()
        //{
        //    if (xlBook.IsInplace)
        //        xlBook.Close(true);
        //}

        //public void CloseApp()
        //{
        //    //TODO: Validar que si este abierto el aplicativo
        //    //TODO: Cerra el libro
        //    //TODO: Liberar Recursos
        //    xlApp.Quit();
        //}

        public void Close()
        {
            xlBook.Close(true);
            xlApp.Quit();

            Process _pro = GetExcelProcess(xlApp);
            _pro.Kill();
        }

        #region AddData

        public void AddDataWithMergue(string _begin, string _end, bool _mergue, int _columnWidth = -1)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.Merge(_mergue);
            if (_columnWidth != -1)
                x.ColumnWidth = _columnWidth;
        }

        public void AddDataString(string _begin, string _end, string _str)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.Value2 = _str;
        }

        public void AddDataInteger(string _begin, string _end, int _int)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.Value2 = _int;
        }

        public void AddDataDouble(string _begin, string _end, double _dbl)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.Value2 = _dbl;
        }

        public void AddDataDateTime(string _begin, string _end, DateTime _date)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.Value2 = _date;
        }

        #endregion AddData

        public void SetVerticalText(string _begin, string _end, string _str)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.Orientation = 90;

            if (!string.IsNullOrEmpty(_str))
            {
                x.Value2 = _str;
            }
        }

        public void SetSizeWith(string _begin, string _end, double _dbl)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.ColumnWidth = _dbl;
        }

        public void SetSizeHigh(string _begin, string _end, double _dbl)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.RowHeight = _dbl;
        }

        public void AddFormat(string _begin, string _end, string _format)
        {
            Range x = xlSheet.get_Range(_begin, _end);
            x.EntireColumn.NumberFormat = _format;
        }

        public void Style(StyleGeneric _style)
        {
            Range x = xlSheet.get_Range(_style.Begin, _style.End);

            if (_style.WrapText != null)
                x.WrapText = _style.WrapText;

            if (_style.Bold != null)
                x.Cells.Font.Bold = _style.Bold;

            if (_style.VerticalAlign == true)
                x.VerticalAlignment = XlHAlign.xlHAlignCenter;

            if (_style.HorizontalAlign == true)
                x.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            if (_style.LineStyle == true)
            {
                x.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDot;
                Borders border = x.Borders;
                border.LineStyle = XlLineStyle.xlContinuous;
                border.Weight = _style.LineWeight;
            }
            if (_style.Color != null)
                x.Interior.Color = System.Drawing.ColorTranslator.ToOle((System.Drawing.Color)_style.Color);
        }

        public void SaveBook()
        {
            xlBook.SaveAs(ruta, XlFileFormat.xlOpenXMLWorkbook);
        }

        #endregion MÉTODOS PÚBLICOS

        #region MÉTODOS PRIVADOS
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        /// <summary>
        /// Cerrado del aplicativo utilizando el Garbage Collector
        /// </summary>
        private void KillTheProcess()
        {
            xlSheet= null;
            xlBook = null;
            xlApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Cerrado del aplicativo utilizando el Garbage Collector
        /// </summary>
        /// <param name="excelApp"></param>
        /// <returns></returns>
        
        private Process GetExcelProcess(Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }
        #endregion
    }
}