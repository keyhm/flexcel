using FlexCel.Core;
using FlexCel.Report;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Example.FlexcelReports.Services
{
    public class FlexCelService
    {
        private const string SUBFOLDER_REPORT_TEMPLATE = "Templates";
        public FlexCelService()
        {
        }

        public DataTable DynamicToStatic(IEnumerable<dynamic> list, string tableName = "tblDatos1")
        {
            var result = new DataTable { TableName = tableName };
            foreach (dynamic item in list)
            {
                var properties = ToDictionary<object>(item);

                var row = result.NewRow();

                foreach (var entry in properties)
                {
                    if (!result.Columns.Contains(entry.Key))
                    {
                        if (entry.Value != null)
                        {
                            result.Columns.Add(entry.Key, entry.Value.GetType());
                        }
                        else
                        {
                            result.Columns.Add(entry.Key);
                        }
                    }
                    row[entry.Key] = entry.Value ?? DBNull.Value;
                }

                result.Rows.Add(row);
            }

            return result;
        }
        public byte[] CreateReportFlexcel(Stream Plantilla, DataSet Datos)
        {
            ExcelFile excelFile;
            XlsFile xlsFile = new XlsFile(true);
            xlsFile.Open(Plantilla);
            using (FlexCelReport flexCelReport = new FlexCelReport())
            {
                foreach (DataTable item in Datos.Tables)
                {
                    flexCelReport.AddTable(item.TableName, item);
                }

                flexCelReport.Run(xlsFile);
                excelFile = xlsFile;
            }

            return ExportarFlexcel(excelFile);
        }
        public Stream GetTemplate(String templateName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "Example.FlexcelReports.Templates." + templateName+ ".xlsx";
            //var documentRelativePath = $"{SUBFOLDER_REPORT_TEMPLATE}\\{templateName}.xlsx";

            var stream = assembly.GetManifestResourceStream(resourceName);

            //var result = Path.GetDirectoryName(Environment.CurrentDirectory);
            //var fullpath =  Path.Combine(result, documentRelativePath);
            //var content = File.OpenRead(path: @"C:\Learning\BackEnd Projects\Example.FlexcelReports\Templates\Informe_Documentos_Eventos_Masivos.xlsx");
            return stream;  
        }

        private IDictionary<string, T> ToDictionary<T>(object source)
        {
            var dictionary = new Dictionary<string, T>();
            foreach (PropertyDescriptor property in TypeDescriptor.GetProperties(source))
            {
                object value = property.GetValue(source);
                if (IsOfType<T>(value) || value == null)
                {
                    dictionary.Add(property.Name, (T)value);
                }
            }
            return dictionary;
        }

        private bool IsOfType<T>(object value)
        {
            return value is T;
        }

        private byte[] ExportarFlexcel(ExcelFile XlsArchivo)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                XlsArchivo.Save(memoryStream, TFileFormats.Xlsx);

                return memoryStream.ToArray();
            }
        }

    }
}
