using Example.FlexcelReports.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Example.FlexcelReports.Services
{
    public class InformeDocumentosEventosMasivosService
    {
        public readonly FlexCelService _flexCelService;
        public InformeDocumentosEventosMasivosService()
        {
            _flexCelService = new FlexCelService();
        }

        public byte[] GenararReporte()
        {
            //** Generar listados y tablas

            var listDocumentos = new List<DocumentoEventoMasivo>() {

                new DocumentoEventoMasivo{ 
                    NumeroDeDocumento = "DOC01", 
                    TipoDeDocumento = "FACTURA-UBL",
                    Proveedor = "Francisco Muñoz",
                    IdentificacionProveedor = "1",
                    Estado = "Radicado",
                    Detalle = "Factura de prueba por excel",
                },
                new DocumentoEventoMasivo{
                    NumeroDeDocumento = "DOC02",
                    TipoDeDocumento = "FACTURA-UBL",
                    Proveedor = "Sebastian Guerrero",
                    IdentificacionProveedor = "1",
                    Estado = "Radicado",
                    Detalle = "Factura de prueba por excel",
                },
                new DocumentoEventoMasivo{
                    NumeroDeDocumento = "DOC03",
                    TipoDeDocumento = "FACTURA-UBL",
                    Proveedor = "Keysy Hernandez",
                    IdentificacionProveedor = "2",
                    Estado = "Radicado",
                    Detalle = "Factura de prueba por excel",
                }
            };

            var listCabecera = new List<CabeceraEventoMasivo>()
            {
                new CabeceraEventoMasivo(){ 
                    Estado = "Completo",
                    FechaGeneracion = DateTime.Now.ToString("yyyy/MM/dd hh-mm-ss"),
                }
            };

            var tablaDocumentos = _flexCelService.DynamicToStatic(listDocumentos, "tblDatos1");

            var tablaCabecera = _flexCelService.DynamicToStatic(listCabecera, "tblDatos2");

            var datosReporte = new DataSet();
            datosReporte.Tables.Add(tablaDocumentos);
            datosReporte.Tables.Add(tablaCabecera);

            var template = "Informe_Documentos_Eventos_Masivos";
            var plantilla = _flexCelService.GetTemplate(template);

            var reporte = _flexCelService.CreateReportFlexcel(plantilla, datosReporte);

            return reporte;
        }
    }
}
