using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace ExcelLibrary.Reader
{
    public class ExcelReader
    {
        #region PROPIEDADES

        public const int ERROR_CONNECTION = 1;
        public const int EXITO = 2;
        public const int HOJA_NO_ENCOTRADA = 3;
        public const int LIBRO_NO_ENCONTRADO = 4;
        private string strPath;
        private OleDbConnection objConnector;
        private OleDbDataAdapter objAdapter;
        private string strNameSheet;
        private int intActionResult;

        #endregion PROPIEDADES

        #region METODOS_PRIVADOS

        private bool Connect()
        {
            try
            {
                if (!System.IO.File.Exists(Path))
                {
                    intActionResult = LIBRO_NO_ENCONTRADO;
                    return false;
                }
                else
                {
                    objConnector = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + Path + ";Extended properties='Excel 12.0 xml;HDR=Yes'");
                    if (objConnector.State.Equals(ConnectionState.Closed))
                    {
                        objConnector.Open();
                    }
                    return true;
                }
            }
            catch (Exception)
            {
                intActionResult = ERROR_CONNECTION;
                Disconnect();
                return false;
            }
        }

        private void Disconnect()
        {
            if (objConnector.State.Equals(ConnectionState.Open))
            {
                objConnector.Close();
            }
        }

        #endregion METODOS_PRIVADOS

        #region METODOS_PUBLICOS

        public ExcelReader()
        {
            strNameSheet = "";
            objConnector = default(OleDbConnection);
            objAdapter = default(OleDbDataAdapter);
        }

        public string Path
        {
            get { return strPath; }
            set { strPath = value; }
        }

        public ActionReturn Load()
        {
            List<List<string>> ListaFila = new List<List<string>>();
            if (Connect())
            {
                if (string.IsNullOrEmpty(strNameSheet))
                {
                    intActionResult = HOJA_NO_ENCOTRADA;
                }
                else
                {
                    try
                    {
                        DataTable dtDatos = new DataTable();
                        objAdapter = new OleDbDataAdapter("select * from [" + strNameSheet + "]", objConnector);
                        objAdapter.Fill(dtDatos);
                        List<string> ListaColumna = new List<string>();
                        foreach (DataColumn column in dtDatos.Columns)
                        {
                            ListaColumna.Add(column.ToString());
                        }
                        ListaFila.Add(ListaColumna);
                        /* fin recorre la primera fila*/
                        foreach (DataRow row in dtDatos.Rows)
                        {
                            ListaColumna = new List<string>();
                            foreach (DataColumn column in dtDatos.Columns)
                            {
                                ListaColumna.Add(row[column].ToString());
                            }
                            ListaFila.Add(ListaColumna);
                        }
                        intActionResult = EXITO;
                    }
                    catch (Exception)
                    {
                        intActionResult = ERROR_CONNECTION;
                    }
                    finally
                    {
                        Disconnect();
                    }
                }
            }
            return new ActionReturn { ListData = ListaFila, Accion = intActionResult };
        }

        public void SheetSearch(int IndexHoja)
        {
            if (Connect())
            {
                DataTable dtHojasNombres = new DataTable();
                dtHojasNombres = objConnector.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string[] excelSheets = new string[dtHojasNombres.Rows.Count];
                int cont = 0;
                foreach (DataRow row in dtHojasNombres.Rows)
                {
                    excelSheets[cont] = row["TABLE_NAME"].ToString();
                    cont++;
                }
                strNameSheet = (excelSheets.Length > IndexHoja) ? excelSheets[IndexHoja] : "";
                Disconnect();
            }
            else
            {
                strNameSheet = "";
            }
        }

        public void SheetSearch(string NombreHoja)
        {
            this.strNameSheet = NombreHoja;
        }

        #endregion METODOS_PUBLICOS
    }
}