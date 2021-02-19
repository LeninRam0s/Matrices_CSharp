/*PARTIENDO DE UNA MATRIZ N*M RELLENAR LA FILA Y LA COLUMA CUANDO ESTA ENCUENTRE UN 0, LA MATRIZ DEBERA SER DINAMICA
 Y SE DEBERA MANEJAR UN ARCHIVO EXCEL PARA LEER LA MATRIZ Y ACTUALIZAR LOS DATOS DE LA NUEVA MATRIZ CON LAS CONDICIONES
OTORGADAS*/


using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Matrices_CSharp.clases
{
    class clsMatrices
    {
        public int rw { get; set; }
        public int cl { get; set; }

        //Recorre la matriz
        public int[,] MatrizXls()
        {
            //ABRIR EXCEL
            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(@"C:\tmp\data.xlsx");
            xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            int[,] datos = new int[rw, cl];

            for (int i = 0; i < rw; i++)
            {
                for (int j = 0; j < cl; j++)
                {
                    datos[i, j] = (int)(range.Cells[i + 1, j + 1] as Excel.Range).Value2;
                }
            }

            xlWorkbook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);

            return datos;
        }

        //POSICION DEL 0 EN LA MATRIZ
        public void Posicion()
        {
            var posicion = MatrizXls();
            for (int i = 0; i < rw; i++)
            {
                for (int j = 0; j < cl; j++)
                {
                    if (posicion[i, j] == 0)
                    {
                        fila = i;
                        colum = j;
                        break;
                    }
                }
            }
        }

        //IMPRIMIR EN CONSOLA
        public int fila { get; set; }
        public int colum { get; set; }
        public int[,] VistaConsola()
        {
            var matriz = MatrizXls();
            Posicion();

            for (int i = 0; i < rw; i++)
            {
                for (int j = 0; j < cl; j++)
                {
                    if (i == fila | j == colum)
                    {
                        matriz[i, j] = 0;
                    }
                    Console.Write(matriz[i, j] + " ");
                }
                Console.WriteLine("");
            }

            fila++;
            colum++;

            Console.WriteLine("\nFila " + fila + ",\nColumna " + colum);//IMPRIME EL VALOR DE LA POSICION
            return matriz;
        }


        //RELLENAR EXCEL
        public void Rellenar()
        {
            //ABRIR EXCEL
            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(@"C:\tmp\data.xlsx");
            xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

            //ESTABLECER LOS RANGOS UTILIZADOS
            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            var datos = VistaConsola();

            for (int i = 0; i < rw; i++)
            {
                for (int j = 0; j < cl; j++)
                {
                    range.Cells[i + 1, j + 1] = datos[i, j];
                }
            }
            Console.WriteLine("\nEjecucion exitosa!");

            xlWorkbook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
