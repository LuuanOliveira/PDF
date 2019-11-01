using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace PdfCofin
{
    class Program
    {
        static void Main(string[] args)
        {
            PdfCofin pdfCofin = new PdfCofin();
            pdfCofin.Page(); //Chamada para o metodo
            Console.WriteLine("teste");
        }
    }
}
