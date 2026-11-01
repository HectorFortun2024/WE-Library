//====================================================================
// CLASE:  WE
// Lenguaje: C#
// FECHA ULTIMA ACTUALIZACION: 30/Oct/2026 (Word con DocumentFormat.OpenXml en lugar de Docx)
//====================================================================
using System;
using System.Drawing;
using System.IO;
using System.Net.Mail;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;                        // Trabajos con Word (SDK de Microsoft, MIT)
using DocumentFormat.OpenXml.Packaging;
using W   = DocumentFormat.OpenXml.Wordprocessing;
using A   = DocumentFormat.OpenXml.Drawing;
using DW  = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
//=====================================================================
//                        N A M E S P A C E
//=====================================================================
// El programador debe incluir un <using LN.LN;> para 
// acceder a estas funciones y agregar una referencia a 
// la DLL de la clase
//=====================================================================

namespace WE
    {
    //================================================================================
    //                                  CLASE LN
    //================================================================================    
    public class WE
        {
        //==============================================================================
        //< ZONA DE DEFINICION DE VARIABLES GLOBALES MIEMBRO DE LN >
        //==============================================================================
        public XLWorkbook      ExcelFile    = new XLWorkbook(); // Define un Libro Excel
        public IXLWorksheet    ExcelSheet   = null;             // Define un Sheet Excel
        WordDoc?               Documento;
		// IXLWorksheet worksheet = workbook.AddWorksheet("NombreDeLaHoja");
        //==============================================================================
        public void MensINFO(string Mensaje)
            {
            MessageBox.Show(Mensaje,
                            "<INFORMATIVO>",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
            }
        //===========================================================================================
        //                                    < METODOS DE EXCEL >
        //===========================================================================================
        //            METODO ExcelFormula() PARA DESCRIBIR UNA FORMULA EN EXCEL
        //===========================================================================================
        // Formato:
        // void ExcelFormula(int CellDef, string Formula)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // Método para definir una formula Excel
        // 
        // 
        // Parámetros de entrada:
        // ValueCell    - Es un string que define el título de la columna.
        // Formula      - Es un string que define la letra que se va a usar ("Arial", "Courrier New", etc.

        // Parámetros de salida: No tiene
        // Retorno:     No tiene (void)
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //-------------------------------------------------------------------------------------------
        public void    ExcelFormula(string CellDef ,string Formula)
                {
                    try
                        {
                        // Establece una fórmula en una "celda"
                        ExcelSheet.Cell(CellDef).FormulaA1 = Formula;
                        }
                    catch (Exception e)
                        {
                        MensINFO("Excepción Detectada en ExcelFormula" + 
                                     "\nCellDef = " + CellDef +
                                     "\nFormula = " + Formula +
                                     "\nSe cancela el proceso" + 
                                     "\n Excepción Detectada: \n" + 
                                     e.ToString());
                        System.Windows.Forms.Application.Exit();
                        }
                }
        //===========================================================================================
        //   Metodo ExcelCreateDocExcel() de Creación de un Documento EXCEL con su Sheet Inicial
        //===========================================================================================
        // Formato:
        // XLWorkbook    ExcelCreateDocExcel(string DocExcel, string SheetName)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método ExcelCreateDocExcel() es usado solamente para crear un nuevo documento EXCEL
        // e inmediatamente después de su ejecución se debe ejecutar el método:
        // XLWorkbook ExcelOpenFile(string PathFile)
        // Esto es obligatorio siempre y cuando se desea ejecutar acciones propias de la definición
        // de todas las características que querramos agregar a dicho documento EXCEL, tales como
        // los títulos de encabezamiento de cada columna, los colores de una celda (Cell), el tipo
        // de letra de una celda, el tamaño, y así sucesivamente.
        // Al término de la ejecución de este método se depositan dos nombres importantes para el
        // trabajo con EXCEL en dos variables globales accesibles por el programador:
        // "ExcelFile"  ->  Contiene el nombre interno a utilizar por los diferentes métodos
        //                  y que esta asociado al documento EXCEL que se está generando.
        // "ExcelSheet" ->  Es el nombre de la hoja (sheet) que contiene la tabla de filas y
        //                  columnas con la que se va a trabajar.
        // Nota Importante:
        // Si a continuación de la ejecución del método ExcelCreateDocExcel se desean usar los 
        // diferentes métodos existentes para la creación de una hoja de cálculo determinada,
        // entonces se debe ejecutar en forma obligatoria al método ExcelOpenFile, caso contrario
        // se produce una excepción de programa.
        // Esto se debe a que al final del código de ExcelCreateDocExcel se ejecuta una instrucción
        // workbook.SaveAs(DocExcel);
        // que realiza la grabación en disco del nuevo documento, cerrándolo y dejando a la hoja
        // Excel vacia, a fín de que posteriormente sea definida bajo la utilización de los distintos
        // métodos que daran el formato deseado. Por tal razón usualmente se debe ejecutar a
        // continuación el correspondiente ExcelOpenFile que abrirá las puertas al programador para
        // poder usar cualquiera de los métodos destinados a esta finalidad.
        // 
        // Parámetros de entrada:
        // 1) string DocExcel   -> string que define el Path y el nombre de fichero que se desea
        //                         Ejemplo: @"D\DocumentosExcel\Alumnos.xlsx"
        // 2) string SheetName  -> string que define el nombre de la hoja Excel (sheet) que vamos a
        //                         llenar con los diferentes métodos definidos con esta finalidad.
        // Parámetros de salida:
        // El nombre interno del documento Excel que se usará en varios parámetros de los métodos
        // definidos. Este nombre es del Tipo de Dato "XLWorkbook".
        //===========================================================================================
        public XLWorkbook    ExcelCreateDocExcel(string DocExcel, string SheetName)
                    {
                    try
                        {
                        var ExcelFile = new XLWorkbook();
                        ExcelSheet = ExcelFile.Worksheets.Add(SheetName);
                        ExcelFile.SaveAs(DocExcel);
                        return ExcelFile;
                        //using   (var workbook = new XLWorkbook())
                        //        {
                        //        var worksheet = workbook.Worksheets.Add(SheetName);
                        //        ExcelFile = workbook;
                        //        ExcelSheet = worksheet;
                        //        workbook.SaveAs(DocExcel);
                        //        return workbook;
                        //        }
                        }
                    catch (Exception e)
                        {
                        MensINFO("Excepción Detectada en ExcelCreateDocExcel()" + 
                                     "\nSe cancela el proceso" + "\n Excepción Detectada: \n" + 
                                     e.ToString());
                        System.Windows.Forms.Application.Exit();
                        return null;
                        }
                    }
        //===========================================================================================
        //                      METODO DE APERTURA DE UN DOCUMENTO EXCEL
        //===========================================================================================
        // Formato:
        // XLWorkbook ExcelOpenFile(string PathFile)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método ExcelOpenFile(string PathFile) tiene el objetivo de abrir el documento en blanco
        // creado anteriormente por el método "ExcelCreateDocExcel". Esto es obligatorio a fin de 
        // evitar un excepción que cancelará su tarea.
        // 
        // Parámetros de entrada:
        // 1) string PathFile   -> string que define el Path y el nombre de fichero creado en el paso
        //                         anterior.
        //                         Ejemplo: @"D\DocumentosExcel\Alumnos.xlsx"
        // Parámetros de salida: No tiene
        // Retorno:
        // El nombre interno del documento Excel que se usará en varios parámetros de los métodos
        // definidos. Este nombre es del Tipo de Dato "XLWorkbook" y es requerido por varios de
        // los métodos a utilizar en la generación del documento.
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //-------------------------------------------------------------------------------------------
        public XLWorkbook ExcelOpenFile(string PathFile)
            {
            try
                {
                ExcelFile = new XLWorkbook(PathFile);
 
                }
            catch(Exception e)
                { 
                MensINFO("Excepción Detectada En Apertura del Documento " + PathFile + 
                             "\nSe cancela el proceso" + "\n Excepción Detectada: \n" + 
                             e.ToString());
                System.Windows.Forms.Application.Exit();
                return null;    
                } 
            return ExcelFile; 
            }
        //===========================================================================================
        //                   METODO ExcelSaveFile() DE SALVA DEL DOCUMENTO EXCEL
        //===========================================================================================
        // Formato:
        // void ExcelSaveFile(string FileName)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método "void ExcelSaveFile(string FileName)", es usado para hacer la grabación del
        // documento generado por programa con el Path y el nombre del fichero dado en el operando
        // string FileName.
        // Nota Importante:
        // Este método finaliza el trabajo con el documento que se genera, por lo que si se desea
        // dentro del mismo programa continuar trabajando, es obligatorio volver a usar el método
        // "ExcelOpenFile", pues en caso contrario se produce una excepción.
        // 
        // Parámetros de entrada:
        // 1) string FileName   -> string que define el Path y el nombre de fichero creado en el paso
        //                         anterior.
        //                         Ejemplo: @"D\DocumentosExcel\Alumnos.xlsx"
        // Parámetros de salida: No tiene
        // Retorno:
        // No tiene (void)
        //-------------------------------------------------------------------------------------------
        public void ExcelSaveFile(string FileName)
            {
            try
                {
                ExcelFile.SaveAs(FileName);
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada En Apertura del Documento " + FileName + 
                             "\nSe cancela el proceso" + "\n Excepción Detectada: \n" + 
                             e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return;
            }
        //===========================================================================================
        //     METODO ExcelFormatTitles() DE ESCRITURA DE UN TITULO COLUMNAR (Row = 1, Column = n)
        //===========================================================================================
        // Formato:
        //void ExcelFormatTitles(int ColumnCell,
        //                      string ValueCell,
        //                      string FontName,
        //                      int FontSize,
        //                      bool FontBold,
        //                      bool FontItalic,
        //                      XLColor BackColor, XLColor ForeColor)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // Los títulos de cada columna definidos por el usuario en EXCEL pueden ser variados y con
        // un formato de estilo definido por el programador. Dado que dichos títulos columnares deben
        // aparecer en la línea 1 (row = 1) pues no es requerido como parámetro y por tal motivo el
        // primero que aparece es un entero que define la columna del título que se desea.
        // 
        // Parámetros de entrada:
        // ColumnCell   - Número de la columna.
        // ValueCell    - Es un string que define el título de la columna.
        // FontName     - Es un string que define la letra que se va a usar ("Arial", "Courrier New", etc.
        // FontSize     - El tamaño de la letra (12, 16, etc.)
        // FontBold     - Variable de tipo bool que si esta en true es que se quiere letra negrita.
        // FontItalic   - Variable de tipo bool que si esta en true es que se quiere letra Itálica.
        // BackColor    - Variable de la clase XLColor que define el color de Background (fondo)
        // ForeColor    - Variable de la clase XLColor que define el color de Foreground (letra)
        // Parámetros de salida: No tiene
        // Retorno:     No tiene (void)
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //-------------------------------------------------------------------------------------------
        public void ExcelFormatTitles(int ColumnCell,
                              string ValueCell,
                              string FontName,
                              int FontSize,
                              bool FontBold,
        bool FontItalic,
        XLColor BackColor, XLColor ForeColor)
        {
            try
                {
                //------------------------------------------------------------------------------------------------
                //                         SET FORMAT PARAMETERS TO THE CELL
                //------------------------------------------------------------------------------------------------
                ExcelSheet.Cell(1, ColumnCell).SetValue(ValueCell).Style.Fill.SetBackgroundColor(BackColor);
                ExcelSheet.Cell(1, ColumnCell).Style.Font.FontColor = ForeColor;
                ExcelSheet.Cell(1, ColumnCell).Style.Font.FontName = FontName;
                ExcelSheet.Cell(1, ColumnCell).Style.Font.FontSize = FontSize;
                ExcelSheet.Cell(1, ColumnCell).Style.Font.Bold = FontBold;
                ExcelSheet.Cell(1, ColumnCell).Style.Font.Italic = FontItalic;
                ExcelSheet.Cell(1, ColumnCell).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ExcelSheet.Column(1).AdjustToContents(); 
                ExcelSheet.Row(1).AdjustToContents();
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada en ExcelFormatTitles()" + 
                             "\nBackColor = " + BackColor.ToString() +
                             "\nForeColor = " + ForeColor.ToString() + 
                             "\nColumnCell = " + ColumnCell.ToString() +
                             "\nValueCell = " + ValueCell +
                             "\nFontName = " + FontName.ToString() +
                             "\nFontBold = " + FontBold.ToString() +
                             "\nFontSize = " + FontSize.ToString() +
                             "\nFontItalic = " + FontItalic.ToString() +
                             "\nSe cancela el Proceso" +  
                             e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return;
            }
        //==================================================================
        //       METODO ExcelFormatCell() DE DAR FORMATO A UNA CELDA
        //==================================================================
        // Formato:
        // void ExcelFormatCell(int RowCell, int ColumnCell,
        //                       string ValueCell,
        //                       string FontName,
        //                       int FontSize,
        //                       bool FontBold,
        //                       bool FontItalic,
        //                       XLColor BackColor, XLColor ForeColor
        //                       )
        //
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método "ExcelFormatCell" es usado para asignar un valor a una celda bajo el formato
        // de (fila, columna) con estilo de color, tipo de letra, tamaño de letra, etc.
        // Parámetros de entrada:
        // TableCell    - Indice de la página con la que se va a trabajar
        // RowCell      - Número de fila.
        // ColumnCell   - Número de columna.
        // ValueCell    - Es un string que define el valor de la columna.
        // FontName     - Es un string que define la letra que se va a usar ("Arial", "Courrier New", etc.
        // FontSize     - El tamaño de la letra (12, 16, etc.)
        // FontBold     - Variable de tipo bool que si esta en true es que se quiere letra negrita.
        // FontItalic   - Variable de tipo bool que si esta en true es que se quiere letra Itálica.
        // BackColor    - Variable de la clase XLColor que define el color de Background de la celda
        // ForeColor    - Variable de la clase XLColor que define el color de Foreground de la celda
        // Parámetros de salida: No tiene
        // Retorno:     No tiene (void)
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //==================================================================
        public void ExcelFormatCell(int RowCell, 
                             int ColumnCell,
                             string ValueCell,
                             string FontName,
                             int FontSize,
                             bool FontBold,
                             bool FontItalic,
        XLColor BackColor, XLColor ForeColor
        )
        {
            try
                {
                //------------------------------------------------------------------------------------------------
                //                         SET FORMAT PARAMETERS TO THE CELL
                //------------------------------------------------------------------------------------------------
                ExcelSheet.Cell(RowCell, ColumnCell).SetValue(ValueCell).Style.Fill.BackgroundColor = BackColor;
                ExcelSheet.Cell(RowCell, ColumnCell).Style.Font.FontColor = ForeColor;
                ExcelSheet.Cell(RowCell, ColumnCell).Style.Font.FontName = FontName;
                ExcelSheet.Cell(RowCell, ColumnCell).Style.Font.FontSize = FontSize;
                ExcelSheet.Cell(RowCell, ColumnCell).Style.Font.Italic = FontItalic;
                ExcelSheet.Cell(RowCell, ColumnCell).Style.Font.Bold = FontBold;
                ExcelSheet.Cell(RowCell, ColumnCell).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada en ExcelFormatCell()" + 
                             "\nRowCell,  = " + RowCell.ToString() +
                             "\nColumnCell = " + ColumnCell.ToString() +
                             "\nBackColor = " + BackColor.ToString() +
                             "\nForeColor = " + ForeColor.ToString() + 
                             "\nFontName = " + FontName.ToString() +
                             "\nFontSize = " + FontSize.ToString() +
                             "\nFontItalic = " + FontItalic.ToString() +
                             "\nFontBold = " + FontBold.ToString() +
                             "\nSe cancela el Proceso" +  
                             e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return;
            }

        //===========================================================================================
        //         METODO ExcelWriteCellRC() DE ESCRITURA DE UNA CELDA DEL DOCUMENTO EXCEL
        //===========================================================================================
        // Formato:
        // void ExcelWriteCellRC(int TableCell,
        //                       int RowCell,
        //                       int ColumnCell,
        //                       string ValueCell)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método "ExcelWriteCellRC" es usado para asignar un valor a una celda bajo el formato
        // de (fila, columna).
        // Parámetros de entrada:
        // TableCell    - Indice de la tabla.
        // RowCell      - Número de fila.
        // ColumnCell   - Número de columna.
        // ValueCell    - Es un string que define el valor de la columna.
        // Parámetros de salida: No tiene
        // Retorno:     No tiene (void)
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //===========================================================================================
        public void ExcelWriteCellRC(int RowCell, int ColumnCell, string ValueCell)
            {
            try
                {
                ExcelSheet.Cell(RowCell, ColumnCell).SetValue(ValueCell).CellRight().DataType.ToString();
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada en ExcelWriteCellRC()" + 
                             "\nRowCell,  = " + RowCell.ToString() +
                             "\nColumnCell = " + ColumnCell.ToString() +
                             "\nColumnCell = " + ValueCell +
                             "\nSe cancela el Proceso" +  
                             e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return;
            }

        //===========================================================================================
        //          METODO ExcelWriteCell() DE ESCRITURA DE UNA CELDA DEL DOCUMENTO EXCEL
        //===========================================================================================
        // Formato:
        // void ExcelWriteCell(string TableCell,
        //                     string InputCell,
        //                     string ValueCell)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método "ExcelWriteCell" es usado para asignar un valor a una celda bajo el formato
        // de (string).
        // Parámetros de entrada:
        // TableCell    - Nombre de la tabla.
        // InputCell    - Celda de salida en formato de string, ejemplo "B2", "C5", etc.
        // ValueCell    - Es un string que define el valor que tendrá la columna.
        // Parámetros de salida: No tiene
        // Retorno:     No tiene (void)
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //===========================================================================================
        public void ExcelWriteCell(string InputCell, string ValueCell)
            {
            try
                {
                ExcelSheet.Cell(InputCell).SetValue(ValueCell).CellRight().DataType.ToString();
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada en ExcelWriteCell()" + 
                             "\nInputCell,  = " + InputCell.ToString() +
                             "\nColumnCell = " + ValueCell +
                             "\nSe cancela el Proceso" +  
                             e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return;
            }

        //===========================================================================================
        //         METODO ExcelGetCellRC() DE ESCRITURA DE UNA CELDA CON NOMBRE EN string
        //===========================================================================================
        // Formato:
        // void ExcelGetCellRC(int    TableCell,
        //                     int    RowCell,
        //                     int    ColumnCell)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método "ExcelGetCellRC" es usado para obtener el valor a una celda bajo el formato (R, C)
        // Parámetros de entrada:
        // TableCell    - Indice de la página del libro.
        // RowCell      - Número de fila de la Celda de entrada.
        // ColumnCell   - Número de columna de la Celda de entrada.
        // Parámetros de salida: No tiene
        // Retorno:     No tiene (void)
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //===========================================================================================

        public string ExcelGetCellRC(int RowCell, int ColumnCell)
            {
            string Celda = "";
            try
                {
                Celda = ExcelSheet.Cell(RowCell, ColumnCell).Value.ToString();
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada en ExcelWriteCell()" + 
                             "\nRowCell,  = " + RowCell.ToString() +
                             "\nColumnCell = " + ColumnCell.ToString() +
                             "\nSe cancela el Proceso" +  
                             e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return Celda;
            }

        //===========================================================================================
        //       METODO ""ExcelGetCell" DE LECTURA DE UNA CELDA DEL DOCUMENTO EXCEL
        //===========================================================================================
        // Formato:
        // string ExcelGetCell(int    TableCell,
        //                     string InputCell)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método "ExcelGetCell" es usado para obtener el valor a una celda bajo el formato string
        // de Excel.
        // Parámetros de entrada:
        // TableCell    - Indice de la página del libro.
        // InputCell    - string que referencia a la Celda de entrada, ej, "A4" , "B2", etc.
        // Parámetros de salida: No tiene
        // Retorno:     Devuelve el valor de la celda en formato string.
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //===========================================================================================
        public string ExcelGetCell(string InputCell)
            {
            string Celda = "";
            try
                {
                Celda = ExcelSheet.Cell(InputCell).Value.ToString();
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada En Lectura de la Celda: " + InputCell + 
                "\nSe Cancela el Proceso\n" + e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return Celda;
            }

        //===========================================================================================
        //     METODO ExcelGetCellDataTypeRC() DE LECTURA DEL TIPO DE DATOS DE UNA CELDA (r, c)
        //===========================================================================================
        // Formato:
        // string ExcelGetCellDataTypeRC(int    TableCell,
        //                               int    RowCell,
        //                               int    ColumnCell)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método "ExcelGetCellDataTypeRC" es usado para obtener el tipo de datos de una celda bajo
        // el formato (r, c) de Excel.
        // Parámetros de entrada:
        // TableCell    - Indice de la página que se desea.
        // RowCell      - int que da el valor de la fila de la Celda de entrada.
        // ColumnCell   - int que da el valor de la columna de la Celda de entrada.
        // Parámetros de salida: No tiene
        // Retorno:     Devuelve el tipo de datos de la celda de Excel en formato string.
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //===========================================================================================
        public string ExcelGetCellDataTypeRC(int RowCell, int ColumnCell)
            {
            string TipoDeDato = "";
            XLDataType Tipo = new ClosedXML.Excel.XLDataType();
            try
                {
                Tipo = ExcelSheet.Cell(RowCell, ColumnCell).DataType;
                TipoDeDato = Tipo.ToString();
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada En Lectura del Tipo de Dato de la Celda: " +
                             "(" + RowCell.ToString() + ", " + ColumnCell.ToString() + ")" +
                             "\nSe Cancela el Programa\n" + e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return TipoDeDato;
            }

        //===========================================================================================
        // METODO ExcelGetCellDataType() DE LECTURA DEL TIPO DE DATO DE UNA CELDA EN FORMATO "string"
        //===========================================================================================
        // Formato:
        // string ExcelGetCellDataType(int    TableCell,
        //                             string InputCell)
        //-------------------------------------------------------------------------------------------
        // Descripción:
        // El método "ExcelGetCellDataType" es usado para obtener el tipo de datos de una celda bajo
        // el formato string de Excel.
        // Parámetros de entrada:
        // TableCell    - Indice de la página que se desea.
        // InputCell    - Dirección de la celda de entrada en formato string de EXCEL ("A4", "B6", etc.)
        // Parámetros de salida: No tiene
        // Retorno:     Devuelve el tipo de datos de la celda de Excel en formato string.
        // Si se detecta algún error en la apertura se emite el mensaje correspondiente con la
        // excepción detectada.
        //===========================================================================================
        public string ExcelGetCellDataType(string InputCell)
            {
            string TipoDeDato = "";
            XLDataType Tipo = new ClosedXML.Excel.XLDataType();
            try
                {
                Tipo = ExcelSheet.Cell(InputCell).DataType;
                TipoDeDato = Tipo.ToString();
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada En Lectura delo Tipo de Dato de la Celda: " + InputCell + 
                             "\nSe Cancela el Programa\n" + e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return TipoDeDato;
            }

        //============================================================================================
        //          METODO ExcelOpenSheet() DE APERTURA DE UNA HOJA (SHEET) DE EXCEL
        //============================================================================================
        public IXLWorksheet ExcelOpenSheet(int SheetNumber)
            {
            try
                {
                ExcelSheet = ExcelFile.Worksheet(SheetNumber);
                }
            catch(Exception e)
                {
                MensINFO("Excepción Detectada en Apertura de una Hoja Excel." + 
                             "\nNúmero de Hoja (Sheet) = " + SheetNumber.ToString() +
                             "\nSe Cancela el Proceso" + 
                             e.ToString());
                }
            return ExcelSheet;
            }
        //========================================================================================================
        //                                      < METODOS WORD >
        //========================================================================================================
        // Desde el 30/Oct/2026 estos métodos usan el SDK oficial de Microsoft "DocumentFormat.OpenXml"
        // (licencia MIT, sin restricciones de uso comercial), en lugar del paquete Docx de Xceed.
        // El documento se maneja con la clase "WE.WordDoc" y las tablas con la clase "WE.WordTable",
        // ambas definidas al final de este fichero. La alineación se da con "WE.Alignment" (left, center,
        // right, both) y el tipo de título con "WE.HeadingType" (Heading1 ... Heading9).
        //========================================================================================================
        // Método de Creación del Documento Word con el Path
        //========================================================================================================
        // Descripción:
        // Es el primer método a ejecutar para la creación de un documento Word.
        // El parámetro PathDoc de tipo string define el Path del documento y el nombre del mismo.
        // Al finalizar la ejecución se retorna el objeto asociado "ObjDoc", que va a ser referenciado
        // a partir de este momento por los restantes métodos asociados.
        // "ObjDoc" es un objeto de la clase "WE.WordDoc".
        // Este objeto retornado se guarda también en la variable global "Documento" del mismo tipo para futuros
        // usos.
        // Nota: El documento se construye en memoria y se escribe en el disco al ejecutar "CloseWordDoc".
        //========================================================================================================
        public WordDoc OpenWordDoc(string PathDoc)
               {
               // Creación de la variable de tipo var ObjDoc
               var ObjDoc = new WordDoc(PathDoc);
               Documento = ObjDoc;
               return ObjDoc;
               }

        //========================================================================================================
        // Método de Creación de un párrafo de tipo Título1, Título2, etc.
        //========================================================================================================
        // Descripción:
        // El método "AddTitleType" de tipo "void" es utilizado para crear un párrafo de tipo Título1, 2, etc.
        // Contiene 9 parámetros:
        // "DocName" de la clase "WE.WordDoc" define el nombre del documento Word retornado en OpenWordDoc.
        // "TitleTXT" de tipo string, define el texto de tipo Título que se va a insertar.
        // "TitleType" es un valor de "WE.HeadingType" que define si el párrafo dado en
        //             "TitleTXT" estará en Título1 (Heading1), Título2 (Heading2), etc.
        // "TitleFontType" de tipo string, define el tipo de letra, por ejemplo "Arial".
        // "TitleFontSize" de tipo int, define el tamaño de la letra, por ejemplo 14.
        // "TitleBoldType" de tipo bool, si es true significa que se desea letra "bold".
        // "TitleItalicType" de tipo bool, si es true significa que se desea letra "Itálica".
        // "System.Drawing.Color TitleColor", define el color de la letra, por ejemplo "System.Drawing.Color.Red"
        //  define color rojo, si escogemos System.Drawing.Color.DarkRed entonces es rojo oscuro, etc.
        // "WE.Alignment TitleAligment", define la alineación del párrafo de título, de modo tal
        // que si escribimos "WE.Alignment.center", entonces el texto aparecerá centrado.
        //========================================================================================================
        public void AddTitleType(WordDoc                        DocName,
                              string                            TitleTXT,
                              HeadingType                       TitleType,
                              string                            TitleFontType,
                              int                               TitleFontSize,
                              bool                              TitleBoldType,
                              bool                              TitleItalicType,
                              System.Drawing.Color              TitleColor,
                              Alignment                         TitleAligment)
            {
            DocName.Agregar(WordDoc.Parrafo(TitleTXT,
                                            WordDoc.Formato(TitleFontType, TitleFontSize, TitleBoldType,
                                                            TitleItalicType, TitleColor),
                                            TitleAligment,
                                            TitleType.ToString()));
            }

        //========================================================================================================
        // Método "CloseWordDoc" de Cierre del Documento Word
        //========================================================================================================
        // Descripción:
        // Es el último método a ejecutar en la creación de un documento Word y es el que crea a dicho documento
        // en el disco.
        // Después de ejecutado el "CloseWordDoc" no se permite el uso de los restantes métodos, a menos que se
        // vuelva a crear otro nuevo documento.
        // El parámetro DocOpen de tipo "WE.WordDoc" define el nombre que fué retornado por el método
        // "OpenWordDoc".
        // Nota: Internamente lo que se ejecuta es el método "Save" de "WE.WordDoc": "ObjDoc.Save();"
        //========================================================================================================
        public void CloseWordDoc(WordDoc DocOpen)
               {
               DocOpen.Save();
               }

        //========================================================================================================
        // Método "InsertImageWordDoc" de Inserción de una Imagen en el Documento Word
        //========================================================================================================
        // Descripción:
        // El método "InsertImageWordDoc", esta destinado a la inserción de una imagen (.jpg, .png, .bmp, .gif,
        // .tif), en la línea en curso, dentro del proceso de creación del documento y requiere de dos parámetros
        // obligatorios:
        // 1) "DocOpen" del tipo "WE.WordDoc" que es el nombre que fué retornado por "OpenWordDoc".
        // 2) El segundo operando se da el path de la imagen de tipo "string".
        // La imagen se inserta con su tamaño original; si es más ancha que el texto de la página, se reduce
        // hasta el ancho del texto conservando su proporción.
        // Es conveniente aclarar que al finalizar se escribe una línea de string "***" para evitar conflictos
        // posteriores en Ashampoo Office Write, que es un equivalente al Word de Microsoft pero más barato. Esto
        // ocurre cuando se crea una tabla inmediatamente después de otra tabla, lo cual provoca que el párrafo
        // siguiente no se visualiza en dicha aplicación, también el WordPad de Windows 10 detecta el problema
        // pero advierte que esto pudiera suceder en apps equivalentes al Word de Microsoft.
        // Este bug, es de suponer que sea arreglado en algun Service Pack de la app.
        //========================================================================================================
        public void InsertImageWordDoc(WordDoc                  DocOpen,
                                       string                   Imagen)
               {
                // Agregar una imagen al documento
                DocOpen.Agregar(new W.Paragraph(new W.Run(DocOpen.CrearImagen(Imagen))));
                AddParagraph(DocOpen, "***", "Arial", 10, true, true, System.Drawing.Color.DarkRed,
                             Alignment.both);
                }

        //========================================================================================================
        // Método "AppendTable" de Creación de una Tabla con #filas y #columnas y sus Bordes.
        //========================================================================================================
        // Descripción:
        // El método "AppendTable" realiza la creación de una tabla vacia en Word, de modo tal que solo tiene en
        // cuenta el número de filas y columnas de la misma. La tabla aparece en el documento cuando se ejecuta
        // "ObjDoc.InsertTable(TABLA)".
        // Este método tiene tres parámetros de entrada:
        // 1) DocName que es el nombre que fué retornado por "OpenWordDoc" del tipo WE.WordDoc
        // 2) TableRows es un parámetro numérico de tipo "int" que define la cantidad de filas de la tabla.
        // 3) TableColumns es un parámetro numérico de tipo "int" que define la cantidad de columnas de la tabla.
        // Al término de la ejecución se retorna el nombre interno de la tabla, el cual es retornado con el tipo de
        // la clase "WE.WordTable" y que es el que debe ser usado posteriormente en todos los métodos
        // asociados al trabajo con tablas.
        // Las columnas se reparten a partes iguales el ancho del texto de la página (se fijan en la rejilla de la
        // tabla para que también las respeten WPS Office y otros equivalentes al Word).
        //========================================================================================================
        public WordTable AppendTable(WordDoc                 DocName,
                                     int                     TableRows,
                                     int                     TableColumns)
               {
               return new WordTable(TableRows, TableColumns, WordDoc.AnchoTexto);
               }

        //========================================================================================================
        // Método "PutHeaderName" de Creación del Header Columnar con letra Bold Standard
        //========================================================================================================
        // Descripción:
        // El método "PutHeaderName" de tipo "void" es el que se ocupa de escribir uno a uno cada título de
        // encabezamiento en cada celda (Columna) de la fila 0 (Fila de encabezado).
        // Contiene 6 parámetros:
        // "TableName"              del tipo (WE.WordTable)
        // "TableCell"              del tipo (int)
        // "ColumnHeader"           del tipo (string)
        // "ColumnHeaderFontType"   del tipo (string)
        // "ColumnHeaderFontSize"   del tipo (int)
        // "ColumnHeaderColor"      del tipo (System.Drawing.Color)
        // Las celdas de las filas se escriben alineadas al centro.
        // En "TableName" aparece el nombre que retornó el método "AppendTable" como nombre de tabla
        // En "TableCell" se da el número de columna del Header Columnar (int).
        // En "ColumnHeader" se define el título de la columna (string).
        // En "ColumnHeaderFontType" define la letra que se va a usar para el ColumnHeader (string).
        // En "ColumnHeaderFontSize" se defne el tamaño de la letra (int).
        // En "ColumnHeaderColor " define el color de la letra (System.Drawing.Color).
        // El número de las columnas (Cells) son relativos a 0 al igual que el de las filas (Rows).
        //========================================================================================================
        public void PutHeaderName(WordTable                 TableName,
                                  int                       TableCell,
                                  string                    ColumnHeader,
                                  string                    ColumnHeaderFontType,
                                  int                       ColumnHeaderFontSize,
                                  System.Drawing.Color      ColumnHeaderColor)
               {
               TableName.Escribir(0, TableCell, ColumnHeader,
                                  WordDoc.Formato(ColumnHeaderFontType, ColumnHeaderFontSize, true, false,
                                                  ColumnHeaderColor),
                                  Alignment.center);
               }

        //========================================================================================================
        // Método "PutHeaderNameItalic" de Creación del Header Columnar con letra Itálica
        //========================================================================================================
        // Descripción:
        // El método "PutHeaderNameItalic" de tipo "void" es el que se ocupa de escribir uno a uno el contenido de
        // cada celda de encabezamiento de la tabla (fila 0), con letra itálica.
        // Contiene 6 parámetros:
        // "TableName"              del tipo (WE.WordTable)
        // "TableCell"              del tipo (int)
        // "ColumnHeader"           del tipo (string)
        // "ColumnHeaderFontType"   del tipo (string)
        // "ColumnHeaderFontSize"   del tipo (int)
        // "ColumnHeaderColor"      del tipo (System.Drawing.Color)
        // Las celdas de las filas se escriben alineadas al centro.
        // En "TableName" aparece el nombre que retornó el método "AppendTable" como nombre de tabla
        // En "TableCell" se da el número de columna del Header Columnar (int).
        // En "ColumnHeader" se define el título de la columna (string).
        // En "ColumnHeaderFontType" define la letra que se va a usar para el ColumnHeader (string).
        // En "ColumnHeaderFontSize" se defne el tamaño de la letra (int).
        // En "ColumnHeaderColor " define el color de la letra (System.Drawing.Color).
        // El número de las columnas (Cells) son relativos a 0 al igual que el de las filas (Rows).
        //========================================================================================================
        public void PutHeaderNameItalic(WordTable                 TableName,
                                        int                       TableCell,
                                        string                    ColumnHeader,
                                        string                    ColumnHeaderFontType,
                                        int                       ColumnHeaderFontSize,
                                        System.Drawing.Color      ColumnHeaderColor)
               {
               TableName.Escribir(0, TableCell, ColumnHeader,
                                  WordDoc.Formato(ColumnHeaderFontType, ColumnHeaderFontSize, true, true,
                                                  ColumnHeaderColor),
                                  Alignment.center);
               }

        //========================================================================================================
        // Método "PutInfoCell" de Escritura del Contenido de un Cell
        //========================================================================================================
        // Descripción:
        // El método "PutInfoCell" de tipo "void" es el que se ocupa de escribir uno a uno el contenido de cada celda
        // en una fila dada.
        // Contiene 4 parámetros:
        // "TableName"              del tipo (WE.WordTable)
        // "TableRowInfo"           del tipo (string)
        // "TableRow"               del tipo (int)
        // "TableCell"              del tipo (int)
        // Las celdas de las filas se escriben alineadas a la izquierda, con letra Arial 12 de color negro.
        // En "TableName" aparece el nombre que se le dió a la tabla mediante la clase "WE.WordTable"
        // En "TableRowInfo" se especifica el contenido que se desea escribir en la celda en formato string.
        // En "TableRow" se da el número de la fila que contiene la celda a escribir.
        // En "TableCell" se da el número de la celda (columna) donde se escribirá el contenido deseado.
        // El número de las columnas (Cells) son relativos a 0 al igual que el de las filas (Rows).
        //========================================================================================================
        public void PutInfoCell(WordTable                   TableName,
                                  string                    TableRowInfo,
                                  int                       TableRow,
                                  int                       TableCell)
               {
               TableName.Escribir(TableRow, TableCell, TableRowInfo,
                                  WordDoc.Formato("Arial", 12, false, false, System.Drawing.Color.Black),
                                  Alignment.left);
               }

        //========================================================================================================
        // Método "CreateDocHeader" de Creación del Encabezamiento del Documento si se Desea (Encabezamiento ó Pie De Página)
        //========================================================================================================
        // Descripción:
        // El método "CreateDocHeader" de tipo "void" solo es utilizado para crear encabezamiento de páginas caso
        // de que se desee. Si se ejecuta más de una vez, el último encabezamiento sustituye al anterior.
        // Contiene 8 parámetros:
        // "DocName" de la clase "WE.WordDoc" define el nombre del documento Word.
        // "HeaderTxt" de tipo string, define el texto de encabezamiento de las páginas.
        // "HeaderFontType" de tipo string, define el tipo de letra a usar, por ejemplo "Arial".
        // "HeaderFontSize" de tipo int, define el tamaño de la letra, por ejemplo 14.
        // "HeaderBoldType" de tipo bool, si es true significa que se desea letra "bold".
        // "HeaderItalicType" de tipo bool, si es true significa que se desea letra "Itálica".
        // "System.Drawing.Color HeaderColor", define el color de la letra, por ejemplo System.Drawing.Color.Red
        //  define color rojo, si escogemos System.Drawing.Color.DarkRed entonces es rojo oscuro.
        // "WE.Alignment HeaderAligment", define la alineación del texto de encabezamiento, por
        // ejemplo, si escribimos "WE.Alignment.center", entonces el texto aparecerá centrado.
        // Nota: Usualmente los colores que se escojen con este fin son atenuados (no dark).
        //========================================================================================================
        public void CreateDocHeader(WordDoc                         DocName,
                                    string                          HeaderTxt,
                                    string                          HeaderFontType,
                                    int                             HeaderFontSize,
                                    bool                            HeaderBoldType,
                                    bool                            HeaderItalicType,
                                    System.Drawing.Color            HeaderColor,
                                    Alignment                       HeaderAligment)
            {
            DocName.CrearEncabezado(WordDoc.Parrafo(HeaderTxt,
                                                    WordDoc.Formato(HeaderFontType, HeaderFontSize, HeaderBoldType,
                                                                    HeaderItalicType, HeaderColor),
                                                    HeaderAligment));
            return;
            }

        //========================================================================================================
        // Método "CreateDocFooter" de Creación del Pié de Página del Documento
        //========================================================================================================
        // Descripción:
        // El método "CreateDocFooter" de tipo "void" solo es utilizado para crear el pié de página caso
        // de que se desee. Si se ejecuta más de una vez, el último pié de página sustituye al anterior.
        // Al texto dado en "FooterTXT" se le añade a continuación el número de cada página (campo PAGE de
        // Word). Hasta el 30/Oct/2026 se añadía el número total de páginas (NUMPAGES).
        // Contiene 8 parámetros:
        // "DocName" de la clase "WE.WordDoc" define el nombre del documento Word.
        // "FooterTXT" de tipo string, define el texto del pié de página.
        // "FooterFontType" de tipo string, define el tipo de letra, por ejemplo "Arial".
        // "FooterFontSize" de tipo int, define el tamaño de la letra, por ejemplo 14.
        // "FooterBoldType" de tipo bool, si es true significa que se desea letra "bold".
        // "FooterItalicType" de tipo bool, si es true significa que se desea letra "Itálica".
        // "System.Drawing.Color FooterColor", define el color de la letra, por ejemplo System.Drawing.Color.Red
        //  define color rojo, si escogemos System.Drawing.Color.DarkRed entonces es rojo oscuro.
        // "WE.Alignment FooterAligment", define la alineación del pié de página, por
        // ejemplo, si escribimos "WE.Alignment.center", entonces el texto aparecerá centrado.
        // Nota: Usualmente los colores que se escojen con este fin son atenuados (no dark).
        //========================================================================================================
        public void CreateDocFooter(WordDoc                         DocName,
                                    string                          FooterTXT,
                                    string                          FooterFontType,
                                    int                             FooterFontSize,
                                    bool                            FooterBoldType,
                                    bool                            FooterItalicType,
                                    System.Drawing.Color            FooterColor,
                                    Alignment                       FooterAligment)
            {
            var footer = WordDoc.Parrafo(FooterTXT,
                                         WordDoc.Formato(FooterFontType, FooterFontSize, FooterBoldType,
                                                         FooterItalicType, FooterColor),
                                         FooterAligment);
            // Campo de Word con el número de la página en curso (Word lo calcula en cada página).
            // Se escribe en 5 fragmentos con el mismo formato que el texto, para que WPS y otros
            // equivalentes al Word muestren el número con la misma letra y tamaño.
            OpenXmlElement[] Partes = { new W.FieldChar { FieldCharType = W.FieldCharValues.Begin },
                                        new W.FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve },
                                        new W.FieldChar { FieldCharType = W.FieldCharValues.Separate },
                                        new W.Text("1"),
                                        new W.FieldChar { FieldCharType = W.FieldCharValues.End } };
            foreach (var Parte in Partes)
                footer.Append(new W.Run(WordDoc.Formato(FooterFontType, FooterFontSize, FooterBoldType,
                                                        FooterItalicType, FooterColor),
                                        Parte));
            DocName.CrearPieDePagina(footer);
            return;
            }

        //========================================================================================================
        // Método "AddParagraph" de Adición de un Párrafo en el Documento
        //========================================================================================================
        // Descripción:
        // El método "AddParagraph" de tipo "void" es utilizado para adicionar un nuevo párrafo al documento.
        // Contiene 8 parámetros:
        // "DocName" de la clase "WE.WordDoc" define el nombre retornado en la creación del documento.
        // "ParagraphTXT" de tipo string, define el texto del párrafo a escribir.
        // "ParagraphFontType" de tipo string, define el tipo de letra a usar en el párrafo, por ejemplo "Arial".
        // "ParagraphFontSize" de tipo int, define el tamaño de la letra, por ejemplo 12.
        // "ParagraphBoldType" de tipo bool, si es true significa que se desea letra "bold".
        // "ParagraphItalicType" de tipo bool, si es true significa que se desea letra "Itálica".
        // "System.Drawing.Color ParagraphColor", define el color de la letra, por ejemplo System.Drawing.Color.Red
        //  define color rojo, si escogemos System.Drawing.Color.DarkRed entonces es rojo oscuro.
        // "WE.Alignment ParagraphAligment", define la alineación del párrafo, por
        // ejemplo, si escribimos "WE.Alignment.center", entonces el texto aparecerá centrado.
        //========================================================================================================
        public void AddParagraph(WordDoc                    DocName,
                                string                      ParagraphTXT,
                                string                      ParagraphFontType,
                                int                         ParagraphFontSize,
                                bool                        ParagraphBoldType,
                                bool                        ParagraphItalicType,
                                System.Drawing.Color        ParagraphColor,
                                Alignment                   ParagraphAligment)
            {
            DocName.Agregar(WordDoc.Parrafo(ParagraphTXT,
                                            WordDoc.Formato(ParagraphFontType, ParagraphFontSize, ParagraphBoldType,
                                                            ParagraphItalicType, ParagraphColor),
                                            ParagraphAligment));
            }


        //=============================================================================================
        //                                  Método DGVGetCell()
        //=============================================================================================
        // FORMATO:
        // string DGVGetCell(DataGridView DGV,
        //                       int DGVRow,
        //                       int DGVColumn)
        //
        // DESCRIPCION:
        // El método DGVGetCell() retorna el contenido de la celda (DGVRow, DGVColumn) en formato
        // de string
        // 
        // Parámetros de entrada:
        // DataGridView DGV
        // int DGVRow
        // int DGVColumn
        //
        // Parámteros de salida: 
        // No tiene
        // Retorno:
        // Retorna el contenido de la celda (DGVRow, DGVColumn) en formato de string.
        // Ejemplo:
        // DGVGetCell(DataGridView1, 4, 1); 
        // Retorna el contenido la celda (4,1) con formato de string.
        //=============================================================================================
        public string DGVGetCell(DataGridView DGV, int DGVRow, int DGVColumn)
            {
            object Temp = null;
            try
                {
                Temp = DGV.Rows[DGVRow].Cells[DGVColumn].Value;
                }
            catch (Exception e)
                {
                MensINFO("Excepción Detectada: " + "\nCelda = (" + DGVRow.ToString() + ", " + DGVColumn.ToString() + ")" +
                         "\nSe Cancela el Proceso\n" + e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return Temp.ToString();
            }

        //=============================================================================================
        //                                  Método DGVPutCell()
        //=============================================================================================
        // FORMATO:
        // void DGVPutCell(DataGridView DGV,
        //                   string CellValue,
        //                   int DGVRow,
        //                   int DGVColumn)
        //
        // DESCRIPCION:
        // El método DGVPutCell() escribe el contenido de CellValue en la celda (DGVRow, DGVColumn)
        //  del DataGridView en formato de string.
        // 
        // Parámetros de entrada:
        // DataGridView DGV
        // string CellValue
        // int DGVRow
        // int DGVColumn
        //
        // Parámteros de salida: 
        // No tiene
        // Retorno:
        // No tiene.
        // Ejemplo:
        // DGVPutCell(DataGridView1, "Irma", 4, 1); 
        // Escribe a la cadena "Irma" en la celda (4,1) en formato de string.
        //=============================================================================================
        public void DGVPutCell(DataGridView DGV, string CellValue, int DGVRow, int DGVColumn)
            {
            try
                {
                DGV.Rows[DGVRow].Cells[DGVColumn].Value = CellValue;
                }
            catch (Exception e)
                {
                MensINFO("Excepción Detectada: " + "\nCelda = (" + DGVRow.ToString() + ", " + DGVColumn.ToString() + ")" +
                         "\nCellValue = " + CellValue +
                         "\nSe Cancela el Proceso\n" + e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return;
            }

        //=============================================================================================
        //                                  Método DGVColumnSumarize()
        //=============================================================================================
        // FORMATO:
        // int DGVColumnSumarize(DataGridView DGV,
        //                       int DGVColumn,
        //                       ref double SUMARIA)
        //
        // DESCRIPCION:
        // El método DGVColumnAverageSelectedRows() realiza el cálculo de la suma de las cantidades
        // que se encuentran en las celdas de una columna dada por DGVColumn del DataGridView del primer
        // operando, dejando el resultado en la variable de tipo double del tercer operando SUMARIA,
        // por referencia.
        // 
        // Parámetros de entrada:
        // DataGridView DGV
        // int DGVColumn
        // ref double SUMARIA
        //
        // Parámteros de salida: 
        // ref double SUMARIA -> Transmisión por referencia de la suma columnar
        // Retorno:
        // Un valor entero que informa como fué la terminación de la ejecución del método.
        // Ejemplo:
        // DGVColumnSumarize(DataGridView1, 4, ref SumaTotal);
        //=============================================================================================
        public int DGVColumnSumarize(DataGridView DGV, int DGVColumn, ref double SUMARIA)
            {
            int i = 0;
            string Temp1 = "";
            object Temp;
            double TempDouble = 0;
            double Suma = 0;
            int Posicion = 0;
            int FILAS = 0;
            FILAS = DGV.Rows.Count;
            try
                {
                for (i = 0; i < FILAS - 1; i++) // Recorrer la colección de Items
                    {
                    Posicion = i;
                    Temp = DGV.Rows[i].Cells[DGVColumn].Value;
                    if (Temp == null)
                        {
                        Temp = null;
                        TempDouble = 0;
                        SUMARIA = 0;
                        return -1;
                        }
                    Temp1 = DGV.Rows[i].Cells[DGVColumn].Value.ToString();
                    TempDouble = double.Parse(Temp1);
                    Temp1 = "";
                    Suma = Suma + TempDouble;
                    }
                SUMARIA = Suma; // Retorno la suma columnar (Sumaria)
                return 0;
                }
            catch
                {
                SUMARIA = 0;
                return -1;
                }
            }

        //=============================================================================================
        //                          Método DGVColumnSumarizeSelectedRows()
        //=============================================================================================
        // FORMATO:
        // int DGVColumnSumarizeSelectedRows(DataGridView DGV,
        //                                  int DGVColumn,
        //                                  ref double SUMARIA)
        //
        // DESCRIPCION:
        // El método DGVColumnAverageSelectedRows() realiza el cálculo de la suma de las cantidades
        // que se encuentran en las celdas seleccionadas de una columna dada por DGVColumn del DataGridView
        // del primer operando, dejando el resultado en la variable de tipo double del tercer operando
        // SUMARIA, por referencia.
        // 
        // Parámetros de entrada:
        // DataGridView DGV
        // int DGVColumn
        // ref double SUMARIA
        //
        // Parámteros de salida: 
        // ref double SUMARIA -> Transmisión por referencia de la suma columnar
        // Retorno:
        // Un valor entero que informa como fué la terminación de la ejecución del método.
        // Ejemplo:
        // DGVColumnSumarizeSelectedRows(DataGridView1, 4, ref SumaTotal);
        //=============================================================================================
        public int DGVColumnSumarizeSelectedRows(DataGridView DGV, int DGVColumn, ref double SUMARIA)
            {
            int     i = 0;
            object  Temp = null;
            double  Suma = 0;
            int     FILAS = 0;
            string  Sel = "";
            FILAS = DGV.SelectedRows.Count - 1;

            try
                {
                //--------------------------------------------------------------------------------------
                DataGridViewSelectedRowCollection Seleccionados = DGV.SelectedRows;
                foreach (DataGridViewRow item in Seleccionados) // Recorrer la colección de Items
                    {
                    Temp = DGV.Rows[i].Cells[DGVColumn].Value;
                    if (Temp == null)
                        {
                        Temp = null;
                        SUMARIA = 0;
                        return -1;
                        }
                    Sel = item.Cells[DGVColumn].Value.ToString().Trim(); // La columna 0 es el TELEFONO
                    Suma = Suma + Double.Parse(Sel);
                    }
                SUMARIA = Suma;
                //--------------------------------------------------------------------------------------
                return 0;
                }
            catch
                {
                SUMARIA = 0;
                return -1;
                }
            }

        //=============================================================================================
        //                                  Método DGVColumnAverage()
        //=============================================================================================
        // FORMATO:
        // public int DGVColumnAverage(DataGridView DGV, 
        //                             int DGVColumn, 
        //                             ref double AVERAGE)
        // DESCRIPCION:
        // El método DGVColumnAverage() realiza el cálculo del promedio de las cantidades
        // que se encuentran en las celdas de una columna dada por DGVColumn del DataGridView del
        // primer operando, dejando el resultado en la variable de tipo double del tercer operando
        // AVERAGE, por referencia.
        // 
        // Parámetros de entrada:
        // DataGridView DGV
        // int DGVColumn
        // ref double AVERAGE
        //
        // Parámteros de salida: 
        // ref double AVERAGE -> Transmisión por referencia
        //
        // Retorno:
        // Un valor entero que informa como fué la terminación de la ejecución del método.
        // Ejemplo:
        // DGVColumnAverage(DataGridView1, 4, ref Promedio);
        //
        //=============================================================================================
        public int DGVColumnAverage(DataGridView DGV, 
                                    int DGVColumn, 
                                    ref double AVERAGE)
            {
            int i = 0;
            string Temp1 = "";
            object Temp = null;
            double TempDouble = 0;
            double Suma = 0;
            int Posicion = 0;
            int FILAS = 0;
            FILAS = DGV.Rows.Count;
            try
                {
                for (i = 0; i < FILAS - 1; i++) // Recorrer la colección de Items
                    {
                    Posicion = i;
                    Temp = DGV.Rows[i].Cells[DGVColumn].Value;
                    if (Temp == null)
                        {
                        Temp = null;
                        TempDouble = 0;
                        AVERAGE = 0;
                        return -1;
                        }
                    Temp1 = DGV.Rows[i].Cells[DGVColumn].Value.ToString();
                    TempDouble = double.Parse(Temp1);
                    Temp1 = "";
                    Suma = Suma + TempDouble;
                    }
                AVERAGE = Suma / (FILAS - 1); // Retorno el la media aritmética
                return 0;
                }
            catch
                {
                AVERAGE = 0;
                return -1;
                }
            }

        //=============================================================================================
        //                          Método DGVColumnAverageSelectedRows()
        //=============================================================================================
        // FORMATO:
        // int DGVColumnAverageSelectedRows(DataGridView DGV,
        //                                  int DGVColumn,
        //                                  ref double AVERAGE)
        //
        // DESCRIPCION:
        // El método DGVColumnAverageSelectedRows() realiza el cálculo del promedio de las cantidades
        // que se encuentran en las celdas seleccionadas de una columna dada por DGVColumn del DataGridView
        // del primer operando, dejando el resultado en la variable de tipo double del tercer operando
        // AVERAGE, por referencia.
        // 
        // Parámetros de entrada:
        // DataGridView DGV
        // int DGVColumn
        // ref double AVERAGE
        //
        // Parámteros de salida: 
        // ref double AVERAGE -> Transmisión por referencia
        // Retorno:
        // Un valor entero que informa como fué la terminación de la ejecución del método.
        // Ejemplo:
        // DGVColumnAverageSelectedRows(DataGridView1, 4, ref Promedio);
        //
        //=============================================================================================
        public int DGVColumnAverageSelectedRows(DataGridView DGV, int DGVColumn, ref double AVERAGE)
            {
            int     i = 0;
            object  Temp = null;
            double  Suma = 0;
            int     FILAS = 0;
            string  Sel = "";
            FILAS = DGV.SelectedRows.Count;

            try
                {
                //--------------------------------------------------------------------------------------
                DataGridViewSelectedRowCollection Seleccionados = DGV.SelectedRows;
                foreach (DataGridViewRow item in Seleccionados) // Recorrer la colección de Items
                    {
                    Temp = DGV.Rows[i].Cells[DGVColumn].Value;
                    if (Temp == null)
                        {
                        Temp = null;
                        AVERAGE = 0;
                        return -1;
                        }
                    Sel = item.Cells[DGVColumn].Value.ToString().Trim(); // La columna 0 es el TELEFONO
                    Suma = Suma + Double.Parse(Sel);
                    }
                AVERAGE = Suma / FILAS;
                //--------------------------------------------------------------------------------------
                return 0;
                }
            catch
                {
                AVERAGE = 0;
                return -1;
                }
            }
	public	bool	CompareFiles(string PathFile1, string PathFile2)
		    {
            try
                {
                byte[] file1Bytes = File.ReadAllBytes(PathFile1);
                byte[] file2Bytes = File.ReadAllBytes(PathFile2);
                
                if 	(file1Bytes.Length == file2Bytes.Length)
		        	{
                    for (int i = 0; i < file1Bytes.Length; i++)
		        		{
                        if 	(file1Bytes[i] != file2Bytes[i])
                            return false;
		        		}
                    return true;
		        	}
                }
            catch (Exception e)
                {
                MensINFO("Excepción Detectada En Comparación de Ficheros: " + e.ToString());
                System.Windows.Forms.Application.Exit();
                }
            return false;
		    } // End CompareFiles

        } // FIN DE LA CLASE LN

    //================================================================================
    //                   TIPOS DE APOYO PARA LOS METODOS DE WORD
    //================================================================================
    // Sustituyen a los tipos del paquete Docx de Xceed (30/Oct/2026):
    //   Xceed.Document.NET.Alignment    -> WE.Alignment
    //   Xceed.Document.NET.HeadingType  -> WE.HeadingType
    //   Xceed.Words.NET.DocX            -> WE.WordDoc
    //   Xceed.Document.NET.Table        -> WE.WordTable
    // Internamente usan el SDK oficial de Microsoft "DocumentFormat.OpenXml" (licencia MIT).
    //================================================================================

    //--------------------------------------------------------------------------------
    // Alineación de un párrafo: izquierda, centro, derecha y justificado (both)
    //--------------------------------------------------------------------------------
    public enum Alignment { left, center, right, both }

    //--------------------------------------------------------------------------------
    // Tipo de título: Título1 (Heading1) ... Título9 (Heading9)
    //--------------------------------------------------------------------------------
    public enum HeadingType { Heading1, Heading2, Heading3, Heading4, Heading5,
                              Heading6, Heading7, Heading8, Heading9 }

    //================================================================================
    //                                  CLASE WordDoc
    //================================================================================
    // Es el documento Word que retorna "OpenWordDoc". Se construye en memoria y se
    // escribe en el disco con "Save()" (lo ejecuta "CloseWordDoc").
    //================================================================================
    public class WordDoc
        {
        //----------------------------------------------------------------------------
        // Página A4 con márgenes de 1 pulgada (igual que la versión anterior con Docx).
        // Las medidas de Word van en twips (1/20 de punto, 1440 por pulgada).
        //----------------------------------------------------------------------------
        internal const int AnchoPagina  = 11906;
        internal const int AltoPagina   = 16838;
        internal const int Margen       = 1440;
        internal const int AnchoTexto   = AnchoPagina - 2 * Margen;
        const long         EmuPorTwip   = 635;    // las imágenes van en EMU (914400 por pulgada)

        public  string                  Ruta;     // Path del documento en el disco
        MemoryStream                    Memoria;
        WordprocessingDocument          Paquete;
        W.Body                          Cuerpo;
        W.SectionProperties             Seccion;
        uint                            ContadorImagenes = 0;
        bool                            Cerrado = false;

        internal WordDoc(string PathDoc)
            {
            Ruta    = PathDoc;
            Memoria = new MemoryStream();
            Paquete = WordprocessingDocument.Create(Memoria, WordprocessingDocumentType.Document);
            var Principal = Paquete.AddMainDocumentPart();
            Seccion = new W.SectionProperties(
                          new W.PageSize   { Width = (UInt32Value)(uint)AnchoPagina, Height = (UInt32Value)(uint)AltoPagina },
                          new W.PageMargin { Top = Margen, Right = (UInt32Value)(uint)Margen,
                                             Bottom = Margen, Left = (UInt32Value)(uint)Margen,
                                             Header = 708U, Footer = 708U, Gutter = 0U });
            Principal.Document = new W.Document(new W.Body(Seccion));
            Cuerpo = Principal.Document.Body!;
            CrearEstilos(Principal);
            }

        //----------------------------------------------------------------------------
        // Inserta en el documento (al final) una tabla creada con "AppendTable"
        //----------------------------------------------------------------------------
        public void InsertTable(WordTable Tabla)
            {
            Agregar(Tabla.Tabla);
            }

        //----------------------------------------------------------------------------
        // Reemplaza un texto por otro en todo el documento (cuerpo, tablas,
        // encabezamientos y pies de página)
        //----------------------------------------------------------------------------
        public void ReplaceText(string TextoViejo, string TextoNuevo)
            {
            if (string.IsNullOrEmpty(TextoViejo))
                return;
            var Raices = new List<OpenXmlElement> { Cuerpo };
            foreach (var Parte in Paquete.MainDocumentPart!.HeaderParts) Raices.Add(Parte.Header);
            foreach (var Parte in Paquete.MainDocumentPart!.FooterParts) Raices.Add(Parte.Footer);
            foreach (var Raiz in Raices)
                foreach (var Parrafo in Raiz.Descendants<W.Paragraph>())
                    {
                    var Textos = Parrafo.Descendants<W.Text>().ToList();
                    bool Hecho = false;
                    // Caso habitual: el texto buscado está dentro de un mismo fragmento
                    foreach (var Texto in Textos)
                        if (Texto.Text.Contains(TextoViejo))
                            {
                            Texto.Text  = Texto.Text.Replace(TextoViejo, TextoNuevo);
                            Texto.Space = SpaceProcessingModeValues.Preserve;
                            Hecho       = true;
                            }
                    if (Hecho || Textos.Count < 2)
                        continue;
                    // Caso raro: repartido entre varios fragmentos -> se une en el primero
                    string Todo = string.Concat(Textos.Select(T => T.Text));
                    if (Todo.Contains(TextoViejo))
                        {
                        Textos[0].Text  = Todo.Replace(TextoViejo, TextoNuevo);
                        Textos[0].Space = SpaceProcessingModeValues.Preserve;
                        foreach (var Texto in Textos.Skip(1))
                            Texto.Text = "";
                        }
                    }
            }

        //----------------------------------------------------------------------------
        // Escribe el documento en el disco (Ruta). Después ya no se puede modificar.
        //----------------------------------------------------------------------------
        public void Save()
            {
            if (Cerrado)
                return;
            Paquete.Dispose();                       // vuelca todas las partes a la memoria
            File.WriteAllBytes(Ruta, Memoria.ToArray());
            Memoria.Dispose();
            Cerrado = true;
            }

        //============================================================================
        //               Métodos internos usados por la clase WE
        //============================================================================
        internal void Agregar(OpenXmlElement Elemento)
            {
            // Las propiedades de la sección (página) deben quedar siempre al final
            Cuerpo.InsertBefore(Elemento, Seccion);
            }

        internal static W.RunProperties Formato(string Fuente, int Tamano, bool Negrita, bool Italica,
                                                System.Drawing.Color Color)
            {
            string MediosPuntos = (Tamano * 2).ToString();
            return new W.RunProperties(
                       new W.RunFonts { Ascii = Fuente, HighAnsi = Fuente, ComplexScript = Fuente, EastAsia = Fuente },
                       new W.Bold   { Val = OnOffValue.FromBoolean(Negrita) },
                       new W.Italic { Val = OnOffValue.FromBoolean(Italica) },
                       new W.Color  { Val = $"{Color.R:X2}{Color.G:X2}{Color.B:X2}" },
                       new W.FontSize { Val = MediosPuntos },
                       new W.FontSizeComplexScript { Val = MediosPuntos });
            }

        internal static W.Justification Justificacion(Alignment Alinear)
            {
            switch (Alinear)
                {
                case Alignment.center: return new W.Justification { Val = W.JustificationValues.Center };
                case Alignment.right:  return new W.Justification { Val = W.JustificationValues.Right };
                case Alignment.both:   return new W.Justification { Val = W.JustificationValues.Both };
                default:               return new W.Justification { Val = W.JustificationValues.Left };
                }
            }

        internal static W.Paragraph Parrafo(string Texto, W.RunProperties FormatoTexto, Alignment Alinear,
                                            string? Estilo = null)
            {
            var Propiedades = new W.ParagraphProperties();
            if (Estilo != null)
                Propiedades.Append(new W.ParagraphStyleId { Val = Estilo });
            Propiedades.Append(Justificacion(Alinear));
            return new W.Paragraph(Propiedades,
                                   new W.Run(FormatoTexto,
                                             new W.Text(Texto) { Space = SpaceProcessingModeValues.Preserve }));
            }

        internal void CrearEncabezado(W.Paragraph Contenido)
            {
            var Principal = Paquete.MainDocumentPart!;
            foreach (var Referencia in Seccion.Elements<W.HeaderReference>().ToList())
                {
                Principal.DeletePart(Referencia.Id!);
                Referencia.Remove();
                }
            var Parte = Principal.AddNewPart<HeaderPart>();
            Parte.Header = new W.Header(Contenido);
            Seccion.PrependChild(new W.HeaderReference { Type = W.HeaderFooterValues.Default,
                                                         Id   = Principal.GetIdOfPart(Parte) });
            }

        internal void CrearPieDePagina(W.Paragraph Contenido)
            {
            var Principal = Paquete.MainDocumentPart!;
            foreach (var Referencia in Seccion.Elements<W.FooterReference>().ToList())
                {
                Principal.DeletePart(Referencia.Id!);
                Referencia.Remove();
                }
            var Parte = Principal.AddNewPart<FooterPart>();
            Parte.Footer = new W.Footer(Contenido);
            Seccion.PrependChild(new W.FooterReference { Type = W.HeaderFooterValues.Default,
                                                         Id   = Principal.GetIdOfPart(Parte) });
            }

        internal W.Drawing CrearImagen(string Imagen)
            {
            ImagePartType Tipo;
            switch (Path.GetExtension(Imagen).ToLowerInvariant())
                {
                case ".png":                 Tipo = ImagePartType.Png;  break;
                case ".bmp":                 Tipo = ImagePartType.Bmp;  break;
                case ".gif":                 Tipo = ImagePartType.Gif;  break;
                case ".tif": case ".tiff":   Tipo = ImagePartType.Tiff; break;
                default:                     Tipo = ImagePartType.Jpeg; break;
                }
            var Principal = Paquete.MainDocumentPart!;
            var Parte     = Principal.AddImagePart(Tipo);
            using (var Fichero = File.OpenRead(Imagen))
                Parte.FeedData(Fichero);

            // Tamaño original de la imagen (según sus píxeles y su resolución)
            long Ancho, Alto;
            using (var Img = System.Drawing.Image.FromFile(Imagen))
                {
                Ancho = (long)(Img.Width  / Img.HorizontalResolution * 914400);
                Alto  = (long)(Img.Height / Img.VerticalResolution   * 914400);
                }
            // Si no cabe en el ancho del texto, se reduce conservando la proporción
            long AnchoMaximo = AnchoTexto * EmuPorTwip;
            if (Ancho > AnchoMaximo)
                {
                Alto  = Alto * AnchoMaximo / Ancho;
                Ancho = AnchoMaximo;
                }

            ContadorImagenes++;
            string Nombre = Path.GetFileName(Imagen);
            return new W.Drawing(
                       new DW.Inline(
                           new DW.Extent { Cx = Ancho, Cy = Alto },
                           new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                           new DW.DocProperties { Id = ContadorImagenes, Name = "Imagen " + ContadorImagenes },
                           new DW.NonVisualGraphicFrameDrawingProperties(
                               new A.GraphicFrameLocks { NoChangeAspect = true }),
                           new A.Graphic(
                               new A.GraphicData(
                                   new PIC.Picture(
                                       new PIC.NonVisualPictureProperties(
                                           new PIC.NonVisualDrawingProperties { Id = 0U, Name = Nombre },
                                           new PIC.NonVisualPictureDrawingProperties()),
                                       new PIC.BlipFill(
                                           new A.Blip { Embed = Principal.GetIdOfPart(Parte) },
                                           new A.Stretch(new A.FillRectangle())),
                                       new PIC.ShapeProperties(
                                           new A.Transform2D(
                                               new A.Offset  { X = 0L, Y = 0L },
                                               new A.Extents { Cx = Ancho, Cy = Alto }),
                                           new A.PresetGeometry(new A.AdjustValueList())
                                               { Preset = A.ShapeTypeValues.Rectangle })))
                                   { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                           { DistanceFromTop = 0U, DistanceFromBottom = 0U,
                             DistanceFromLeft = 0U, DistanceFromRight = 0U });
            }

        //----------------------------------------------------------------------------
        // Estilos del documento: Normal y Título1 ... Título9 (Heading1 ... Heading9).
        // La letra, el tamaño y el color los pone cada método, por eso los títulos solo
        // definen el nivel de esquema (para índices y el panel de navegación).
        //----------------------------------------------------------------------------
        static void CrearEstilos(MainDocumentPart Principal)
            {
            var Estilos = new W.Styles(
                new W.DocDefaults(
                    new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
                        new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri",
                                         ComplexScript = "Calibri", EastAsia = "Calibri" },
                        new W.FontSize { Val = "22" },
                        new W.FontSizeComplexScript { Val = "22" })),
                    new W.ParagraphPropertiesDefault(new W.ParagraphPropertiesBaseStyle(
                        // Interlineado sencillo y sin espacio entre párrafos (igual que con Docx)
                        new W.SpacingBetweenLines { After = "0", Line = "240",
                                                    LineRule = W.LineSpacingRuleValues.Auto }))),
                new W.Style(new W.StyleName { Val = "Normal" }, new W.PrimaryStyle())
                    { Type = W.StyleValues.Paragraph, StyleId = "Normal", Default = true });
            for (int Nivel = 1; Nivel <= 9; Nivel++)
                Estilos.Append(new W.Style(
                    new W.StyleName { Val = "heading " + Nivel },
                    new W.BasedOn { Val = "Normal" },
                    new W.NextParagraphStyle { Val = "Normal" },
                    new W.PrimaryStyle(),
                    new W.StyleParagraphProperties(
                        new W.KeepNext(),
                        new W.SpacingBetweenLines { Before = Nivel == 1 ? "480" : "200", After = "0" },
                        new W.OutlineLevel { Val = Nivel - 1 }))
                    { Type = W.StyleValues.Paragraph, StyleId = "Heading" + Nivel });
            var Parte = Principal.AddNewPart<StyleDefinitionsPart>();
            Parte.Styles = Estilos;
            }
        } // FIN DE LA CLASE WordDoc

    //================================================================================
    //                                  CLASE WordTable
    //================================================================================
    // Es la tabla que retorna "AppendTable". Se rellena con PutHeaderName,
    // PutHeaderNameItalic y PutInfoCell, y se coloca en el documento con
    // "ObjDoc.InsertTable(TABLA)".
    //================================================================================
    public class WordTable
        {
        internal W.Table Tabla;
        public   int     Rows;
        public   int     Columns;

        internal WordTable(int Filas, int Columnas, int AnchoTotal)
            {
            Rows    = Filas;
            Columns = Columnas;
            string AnchoColumna = (AnchoTotal / Columnas).ToString();

            // Borde triple negro de 1.5 puntos, igual que la versión anterior
            W.TableBorders Bordes = new W.TableBorders(
                new W.TopBorder              { Val = W.BorderValues.Triple, Size = 12U, Space = 0U, Color = "000000" },
                new W.LeftBorder             { Val = W.BorderValues.Triple, Size = 12U, Space = 0U, Color = "000000" },
                new W.BottomBorder           { Val = W.BorderValues.Triple, Size = 12U, Space = 0U, Color = "000000" },
                new W.RightBorder            { Val = W.BorderValues.Triple, Size = 12U, Space = 0U, Color = "000000" },
                new W.InsideHorizontalBorder { Val = W.BorderValues.Triple, Size = 12U, Space = 0U, Color = "000000" },
                new W.InsideVerticalBorder   { Val = W.BorderValues.Triple, Size = 12U, Space = 0U, Color = "000000" });

            Tabla = new W.Table(new W.TableProperties(
                                    new W.TableWidth  { Width = AnchoTotal.ToString(), Type = W.TableWidthUnitValues.Dxa },
                                    Bordes,
                                    new W.TableLayout { Type = W.TableLayoutValues.Fixed }));
            // Rejilla de columnas: WPS Office y otros solo respetan el ancho si está aquí
            var Rejilla = new W.TableGrid();
            for (int C = 0; C < Columnas; C++)
                Rejilla.Append(new W.GridColumn { Width = AnchoColumna });
            Tabla.Append(Rejilla);
            for (int F = 0; F < Filas; F++)
                {
                var Fila = new W.TableRow();
                for (int C = 0; C < Columnas; C++)
                    Fila.Append(new W.TableCell(
                                    new W.TableCellProperties(
                                        new W.TableCellWidth { Width = AnchoColumna, Type = W.TableWidthUnitValues.Dxa }),
                                    new W.Paragraph()));
                Tabla.Append(Fila);
                }
            }

        //----------------------------------------------------------------------------
        // Añade un texto con su formato a la celda (Fila, Columna), relativas a 0
        //----------------------------------------------------------------------------
        internal void Escribir(int Fila, int Columna, string Texto, W.RunProperties Formato, Alignment Alinear)
            {
            var Parrafo = Tabla.Elements<W.TableRow>().ElementAt(Fila)
                               .Elements<W.TableCell>().ElementAt(Columna)
                               .Elements<W.Paragraph>().First();
            if (Parrafo.ParagraphProperties == null)
                Parrafo.PrependChild(new W.ParagraphProperties());
            Parrafo.ParagraphProperties!.Justification = WordDoc.Justificacion(Alinear);
            Parrafo.Append(new W.Run(Formato, new W.Text(Texto) { Space = SpaceProcessingModeValues.Preserve }));
            }
        } // FIN DE LA CLASE WordTable
    } // FIN DEL NAMESPACE LN

