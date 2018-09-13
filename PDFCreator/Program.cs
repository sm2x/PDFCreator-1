using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            PDFGenerator gen = new PDFGenerator();
            gen.SpoolPDF();
        }                  
    }
}
