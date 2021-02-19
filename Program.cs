using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Matrices_CSharp.clases;

namespace Matrices_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            var matriz = new clsMatrices();
            matriz.Rellenar();
            Console.ReadKey();
        }
    }
}
