using Example.FlexcelReports.Services;
using System;
using System.IO;

namespace Example.FlexCelReports
{
    class Program
    {
        static void Main(string[] args)
        {
            var reporte = new InformeDocumentosEventosMasivosService().GenararReporte();

            var str = System.Text.Encoding.Default.GetString(reporte);

            Console.WriteLine(str);
        }
    }

}