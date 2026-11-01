using System.Windows.Forms;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using LN;
using WE;
namespace PruebaWord
    {
    public partial class Form1 : Form
        {
        //----------------------------------------------------
        // Area de definición de variables globales a la clase
        //----------------------------------------------------
        LN.LN OBJ = new LN.LN();
        WE.WE WEOBJ  = new WE.WE();
        //WE.WordDoc           Documento; // Clase WordDoc usada
        //                                  // para definir a
        //                                  // Documento
        //----------------------------------------------------
        public Form1()
            {
            InitializeComponent();
            }
        //WEOBJ.OpenWordDoc(@"D:\DocWord.docx");
        private void button1_Click(object sender, EventArgs e)
            {
            var ObjDoc = WEOBJ.OpenWordDoc(@"D:\AAA.docx");
            //Documento = ObjDoc; // Esto es si se necesita guardar a ObjDoc en un buffer Documento
            // para alguna acción posterior. En nuestro caso no es necesario
            //==========================<CREACION DE LOS HEADERS>=======================
            WEOBJ.CreateDocHeader(ObjDoc, "Encabezado del Documento", "Arial", 14, true, true, System.Drawing.Color.Gray,
                                WE.Alignment.right);
            //==========================<CREACION DE LOS HEADERS>=======================
            WEOBJ.CreateDocHeader(ObjDoc, "Encabezado del Documento", "Arial", 14, true, true, System.Drawing.Color.Gray,
                                WE.Alignment.right);
            //========================<CREACION DE LOS FOOTERS>=========================
            WEOBJ.CreateDocFooter(ObjDoc, "Página # ", "Arial", 14, true, true, System.Drawing.Color.DarkGray,
                                WE.Alignment.right);
            //======================< CREACION DEL PARRAFO1>============================
            WEOBJ.AddTitleType(ObjDoc, "PARRAFO1", WE.HeadingType.Heading1, "Arial", 16, false, true, 
                               System.Drawing.Color.Black,
                             WE.Alignment.both);
            WEOBJ.AddParagraph(ObjDoc, "Este es el texto demostrativo del primer párrafo de prueba de creación de un documento Word mediante programa. " +
                         "Aquí pretendemos mostrar la relativa facilidad con que se crea un documento Word con este conjunto de métodos.",
                         "Arial",
                         10, false, true, System.Drawing.Color.DarkRed, WE.Alignment.both);
            WEOBJ.AddParagraph(ObjDoc, "Este es el segundo párrafo demostrativo de este programa de ejemplo, es conveniente aclarar que nuestro objetivo " +
                         "principal es simplificar al máximo la programación de aplicaciones orientadas al Word.",
                         "Arial",
                         10, false, false, System.Drawing.Color.DarkGreen, WE.Alignment.both);
            //-----------------------------------------------------------------------------------------------------------
            var TABLA = WEOBJ.AppendTable(ObjDoc, 3, 3);
            //-----------------------------------------------------------------------------------------------------------
            WEOBJ.PutHeaderName(TABLA, 0, "NOMBRE", "Arial", 14, System.Drawing.Color.Black);
            WEOBJ.PutHeaderName(TABLA, 1, "EDAD", "Times New Roman", 14, System.Drawing.Color.DarkRed);
            WEOBJ.PutHeaderName(TABLA, 2, "SEXO", "Courrier New", 14, System.Drawing.Color.DarkMagenta);
            //-----------------------------------------------------------------------------------------------------------
            WEOBJ.PutInfoCell(TABLA, "HECTOR",1,0);
            WEOBJ.PutInfoCell(TABLA, "75",1,1);
            WEOBJ.PutInfoCell(TABLA, "M",1, 2);
            WEOBJ.PutInfoCell(TABLA, "IRMA",2,0);
            WEOBJ.PutInfoCell(TABLA, "60",2,1);
            WEOBJ.PutInfoCell(TABLA, "F",2, 2);
            //-----------------------------------------------------------------------------------------------------------
            ObjDoc.InsertTable(TABLA);
            //-----------------------------------------------------------------------------------------------------------
            WEOBJ.AddParagraph(ObjDoc, "AAA.", "Arial", 10, false,  true, System.Drawing.Color.DarkRed,
                         WE.Alignment.both);
            WEOBJ.AddParagraph(ObjDoc, "BBB.", "Arial", 10, false,  false, System.Drawing.Color.DarkRed,
                         WE.Alignment.both);
            //----------------------------------------------------------------------------------------------------------
            var TABLA1 = WEOBJ.AppendTable(ObjDoc, 4, 3);
            //----------------------------------------------------------------------------------------------------------
            WEOBJ.PutInfoCell(TABLA1, "Elvis",1,0);
            WEOBJ.PutInfoCell(TABLA1, "17",1,1);
            WEOBJ.PutInfoCell(TABLA1, "M",1, 2);
            WEOBJ.PutInfoCell(TABLA1, "Gretel",2,0);
            WEOBJ.PutInfoCell(TABLA1, "40",2,1);
            WEOBJ.PutInfoCell(TABLA1, "F",2, 2);
            WEOBJ.PutInfoCell(TABLA1, "Brandon",3,0);
            WEOBJ.PutInfoCell(TABLA1, "1",3,1);
            WEOBJ.PutInfoCell(TABLA1, "M",3, 2);
            //-----------------------------------------------------------------------------------------------------------
            WEOBJ.PutHeaderNameItalic(TABLA1, 0, "NOMBRE", "Arial", 18, System.Drawing.Color.Red);
            WEOBJ.PutHeaderNameItalic(TABLA1, 1, "EDAD", "Arial", 18, System.Drawing.Color.Red);
            WEOBJ.PutHeaderNameItalic(TABLA1, 2, "SEXO", "Arial", 18, System.Drawing.Color.Red);
            ObjDoc.InsertTable(TABLA1);
            //-----------------------------------------------------------------------------------------------------------
            WEOBJ.AddParagraph(ObjDoc, "XXX.", "Arial", 10, false, true, System.Drawing.Color.DarkRed,
                         	   WE.Alignment.both);
            WEOBJ.AddParagraph(ObjDoc, @"https://www.sunpass.com/", "Arial", 10, false,  true, 
							 System.Drawing.Color.DarkRed, WE.Alignment.left);
            //-----------------------------------------------------------------------------------------------------------
            WEOBJ.InsertImageWordDoc(ObjDoc, @"D:\PAISAJE.jpg");
            WEOBJ.AddParagraph(ObjDoc, "XXX.", "Arial", 10, false, true, System.Drawing.Color.DarkRed,
                             WE.Alignment.both);
            ObjDoc.ReplaceText("AAA.", "Secuencia de reemplazo de AAA.");
            ObjDoc.ReplaceText("BBB.", "Secuencia de reemplazo de BBB.");
            ObjDoc.ReplaceText("XXX.", "Señor XXX, de ahora en adelante se llamará YYY.");
            WEOBJ.CloseWordDoc(ObjDoc);
            OBJ.MensINFO(@"Terminación Exitosa del Programa - Resultado en D:\\AAA.docx");
            Application.Exit(); // Fin del Proceso
            }
        }
    }
