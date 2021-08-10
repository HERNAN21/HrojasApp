using DevExpress.AspNetCore.Spreadsheet;
using DevExpress.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Data;

using System.Drawing;

using HrojasApp.Helpers;
using HrojasApp.DTO;
using HrojasApp.Util;
using HrojasApp.Providers;

namespace HrojasApp.Controllers
{
    [ApiExplorerSettings(IgnoreApi = true)]
    [Controller]
    [Route("[controller]/[Action]")]
    public class ToolsController : Controller
    {
        
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        [HttpPost]
        public IActionResult SpreadSheet(DocumentContentFromBytes doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (doc.DocumentId != null) doc.IsEditable = true;
            if (doc.Base64.Contains(",", StringComparison.InvariantCulture))
                doc.Base64 = doc.Base64.Split(',').Last();
            if (!doc.FileName.Contains(".xlsx", StringComparison.InvariantCulture))
                doc.FileName = $"{doc.FileName}_{AppDateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture)}.xlsx";
            if (doc.ContentAccessor == null && doc.Base64 != null)
            {
                doc.DocumentId = doc.FileName;
                doc.ContentAccessor = () => Convert.FromBase64String(doc.Base64);
            }
            return View(doc);
        }

        [HttpPost]
        [HttpGet]
        public IActionResult DxDocRequest()
        {
            return SpreadsheetRequestProcessor.GetResponse(HttpContext);
        }

        [HttpPost]
        public IActionResult Download(SpreadsheetClientState spreadsheetState)
        {
            var spreadsheet = SpreadsheetRequestProcessor.GetSpreadsheetFromState(spreadsheetState);
            var filename = spreadsheet.DocumentId.Split('\\').Last();
            MemoryStream stream = new MemoryStream();
            spreadsheet.SaveCopy(stream, DocumentFormat.Xlsx);
            stream.Position = 0;
            return File(stream, XlsxContentType, filename);
        }

        
        [HttpGet]
        public IActionResult Get()
        {
            List<ConsultarTransferenciasRDTO> lista = new List<ConsultarTransferenciasRDTO>();

            lista.Add(new ConsultarTransferenciasRDTO
                {codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 211, nroOperacion = 15, codAccionista = "00001000", desAccionista = "WEYDERT MENDOZA ZAIDA DEL PILAR", txtOperacion = "VENTA ", compra = 0, venta = 6600, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-25T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 1, codCtaAccionista = "98", nroCertificado = null, codSerieCertificado = null, idCartera = "001001"}
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 211, nroOperacion = 15, codAccionista = "00002615", desAccionista = "HORRUITINER IZQUIERDO LUIS FELIPE", txtOperacion = "VENTA ", compra = 0, venta = 6600, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-25T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 2, codCtaAccionista = "98", nroCertificado = null, codSerieCertificado = null, idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 211, nroOperacion = 15, codAccionista = "00000609", desAccionista = "GARATE CAMACHO GUSTAVO ONOFRE", txtOperacion = "COMPRA", compra = 13200, venta = 0, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-25T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 3, codCtaAccionista = "99", nroCertificado = null, codSerieCertificado = null, idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 210, nroOperacion = 25, codAccionista = "00002603", desAccionista = "FONDO DE VIVIENDA POLICIAL-FOVIPOL", txtOperacion = "VENTA ", compra = 0, venta = 43800, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-26T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 1, codCtaAccionista = "98", nroCertificado = null, codSerieCertificado = null, idCartera = "001001" }
            );
            
            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 210, nroOperacion = 25, codAccionista = "00000609", desAccionista = "GARATE CAMACHO GUSTAVO ONOFRE", txtOperacion = "COMPRA", compra = 43800, venta = 0, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-26T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 2, codCtaAccionista = "99", nroCertificado = null, codSerieCertificado = null, idCartera = "001001" }
            );
            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 212, nroOperacion = 5, codAccionista = "00000127", desAccionista = "CRUZ ROJAS TATIANA ROSA", txtOperacion = "VENTA ", compra = 0, venta = 14400, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-26T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 1, codCtaAccionista = "98", nroCertificado = null, codSerieCertificado = null, idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 212, nroOperacion = 5, codAccionista = "00002379", desAccionista = "EGOAVIL TORRES ERNESTO JESUS", txtOperacion = "VENTA ", compra = 0, venta = 3800, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-26T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 2, codCtaAccionista = "98", nroCertificado = null, codSerieCertificado = null, idCartera = "001001" }
);

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 212, nroOperacion = 5, codAccionista = "00002400", desAccionista = "SAHUANAY RODRIGUEZ LAUREANO", txtOperacion = "VENTA ", compra = 0, venta = 21800, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-26T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 3, codCtaAccionista = "98", nroCertificado = null, codSerieCertificado = null, idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 212, nroOperacion = 5, codAccionista = "00000609", desAccionista = "GARATE CAMACHO GUSTAVO ONOFRE", txtOperacion = "COMPRA", compra = 40000, venta = 0, codAgenteCompra = "0047", agenteCompra = "CAR", codAgenteVende = "0047", agenteVende = "CAR", fecOperacion = "2021-05-26T00:00:00", codNaturaleza = "03", desNaturaleza = "OPER.DE REPORTE A PLAZO CULMINADA", nroSecuencia = 4, codCtaAccionista = "99", nroCertificado = null, codSerieCertificado = null, idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 298, nroOperacion = 1, codAccionista = "00001578", desAccionista = "ALARCON FASANANDO JESSENIA LUZ", txtOperacion = "VENTA ", compra = 0, venta = 3, codAgenteCompra = "0000", agenteCompra = "FPN", codAgenteVende = "0000", agenteVende = "FPN", fecOperacion = "2021-06-02T00:00:00", codNaturaleza = "30", desNaturaleza = "TRANSFERENCIA DE ACCIONES", nroSecuencia = 1, codCtaAccionista = "00", nroCertificado = "00000005", codSerieCertificado = null, idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 298, nroOperacion = 1, codAccionista = "00001571", desAccionista = "ABUSADA ABEDRABO NATTY", txtOperacion = "COMPRA", compra = 3, venta = 0, codAgenteCompra = "0000", agenteCompra = "FPN", codAgenteVende = "0000", agenteVende = "FPN", fecOperacion = "2021-06-02T00:00:00", codNaturaleza = "30", desNaturaleza = "TRANSFERENCIA DE ACCIONES", nroSecuencia = 2, codCtaAccionista = "00", nroCertificado = "00000520", codSerieCertificado = "000", idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 238, nroOperacion = 1, codAccionista = "00001576", desAccionista = "AGUILAR PINILLOS ANGEL REMIGIO", txtOperacion = "VENTA ", compra = 0, venta = 137, codAgenteCompra = "0000", agenteCompra = "FPN", codAgenteVende = "0000", agenteVende = "FPN", fecOperacion = "2021-06-29T00:00:00", codNaturaleza = "32", desNaturaleza = "REDUCCION DE CAPITAL SOCIAL", nroSecuencia = 1, codCtaAccionista = "00", nroCertificado = "00000004", codSerieCertificado = null, idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 238, nroOperacion = 1, codAccionista = "00001594", desAccionista = "ALVARADO ESPINOZA MONICA MERCEDES", txtOperacion = "VENTA ", compra = 0, venta = 3, codAgenteCompra = "0000", agenteCompra = "FPN", codAgenteVende = "0000", agenteVende = "FPN", fecOperacion = "2021-06-29T00:00:00", codNaturaleza = "32", desNaturaleza = "REDUCCION DE CAPITAL SOCIAL", nroSecuencia = 2, codCtaAccionista = "00", nroCertificado = "00000015", codSerieCertificado = null, idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 238, nroOperacion = 1, codAccionista = "00001576", desAccionista = "AGUILAR PINILLOS ANGEL REMIGIO", txtOperacion = "COMPRA", compra = 100, venta = 0, codAgenteCompra = "0000", agenteCompra = "FPN", codAgenteVende = "0000", agenteVende = "FPN", fecOperacion = "2021-06-29T00:00:00", codNaturaleza = "32", desNaturaleza = "REDUCCION DE CAPITAL SOCIAL", nroSecuencia = 3, codCtaAccionista = "00", nroCertificado = "00000518", codSerieCertificado = "001", idCartera = "001001" }
            );

            lista.Add(new ConsultarTransferenciasRDTO
                { codEmpresa = "001", nomEmpresa = "ENEL DISTRIBUCIÓN PERÚ S.A.A.", codCartera = "001", nomCartera = "ENEL DX PERU - ACCIONES COMUNES", nroActa = 238, nroOperacion = 1, codAccionista = "00001594", desAccionista = "ALVARADO ESPINOZA MONICA MERCEDES", txtOperacion = "COMPRA", compra = 1, venta = 0, codAgenteCompra = "0000", agenteCompra = "FPN", codAgenteVende = "0000", agenteVende = "FPN", fecOperacion = "2021-06-29T00:00:00", codNaturaleza = "32", desNaturaleza = "REDUCCION DE CAPITAL SOCIAL", nroSecuencia = 4, codCtaAccionista = "00", nroCertificado = "00000519", codSerieCertificado = "001", idCartera = "001001" }
            );

            Workbook workbook = new Workbook();

            var data = lista;

            var dtEmpresaCartera = (from a in data
                                    select new
                                    {
                                        codEmp = a.codEmpresa,
                                        codCart = a.codCartera,
                                        nomCartera = a.nomCartera,
                                        nomEmpresa = a.nomEmpresa
                                    }).Distinct();


            foreach (var item in dtEmpresaCartera)
            {
                string nombreHoja = string.Empty;

                if (item.nomCartera.Trim().ToString().Contains("BUENAVENTURA"))
                {
                    nombreHoja = item.nomCartera.Trim().ToString().Replace("BUENAVENTURA", "").Replace("-", "");
                }
                else
                {
                    nombreHoja = $"{ item.codEmp }-{ item.codCart }";

                }

                workbook.Worksheets.Add().Name = nombreHoja;

                Worksheet worksheet = workbook.Worksheets[nombreHoja];


                CellRange cellA1 = worksheet["A1:L1"];
                cellA1.Value = item.nomEmpresa;
                cellA1.Merge();
                cellA1.Font.Bold = true;
                cellA1.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                CellRange cellA2 = worksheet["A2:L2"];
                cellA2.Value = "TRANSFERENCIA DE ACCIONES";
                cellA2.Merge();
                cellA2.Font.Bold = true;
                cellA2.Font.Color = Color.OrangeRed;
                cellA2.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                CellRange cellA3 = worksheet["A3:L3"];
                cellA3.Merge();
                cellA3.Value = "DEL";
                cellA3.Font.Bold = true;
                cellA3.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                CellRange cellA4 = worksheet["A4:L4"]; cellA4.Merge();
                cellA4.Value = item.nomCartera.Trim().ToString();
                cellA4.Font.Bold = true;
                cellA4.Font.Color = Color.DodgerBlue;
                cellA4.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                var dtResumen = (from a in data
                                 where a.codEmpresa == item.codEmp && a.codCartera == item.codCart
                                 select new
                                 {
                                     FEC_OPERACION = a.fecOperacion,
                                     NRO_ACTA = a.nroActa,
                                     NRO_OPERACION = a.nroOperacion,
                                     COD_ACCIONISTA = a.codAccionista,
                                     DES_ACCIONISTA = a.desAccionista,
                                     COD_CUENTA_ACCIONISTA = a.codCtaAccionista,
                                     DES_NATURALEZA = a.desNaturaleza,
                                     AGENTE_COMPRA = a.agenteCompra,
                                     COMPRA = a.compra,
                                     AGENTE_VENDE = a.agenteVende,
                                     VENTA = a.venta,
                                     NRO_CERTIFICADO = a.nroCertificado
                                 }).OrderBy(a => a.DES_NATURALEZA).ThenBy(a => a.FEC_OPERACION).ThenByDescending(a => a.COMPRA).ThenByDescending(a => a.VENTA).ToList();
                DataTable dtDatosCartera = HelperDatos.ToDataTable(dtResumen);
                worksheet.Import(dtDatosCartera, true, 4, 0);


                Table table = worksheet.Tables.Add(worksheet["A5:L" + (dtDatosCartera.Rows.Count + 11) + ""], true);
                table.HeaderRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                table.HeaderRowRange.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                table.HeaderRowRange.Alignment.WrapText = true;
                table.HeaderRowRange.RowHeight = 120;

                table.Style = workbook.TableStyles[BuiltInTableStyleId.TableStyleLight9];

                TableColumn ColumnFecha = table.Columns[0];//0
                ColumnFecha.Name = "FECHA";

                TableColumn ColumnNroActa = table.Columns[1];
                ColumnNroActa.Name = "NRO. ACTA";

                TableColumn ColumnNroOperacion = table.Columns[2];
                ColumnNroOperacion.Name = "NRO. OPERACION";

                TableColumn ColumnCodAcc = table.Columns[3];//1
                ColumnCodAcc.Name = "CÓD. ACCIONISTA";
                TableColumn ColumnDescripcion = table.Columns[4];//2
                ColumnDescripcion.Name = "DESCRIPCIÓN";

                TableColumn ColumnCta = table.Columns[5];
                ColumnCta.Name = "COD. CUENTA";

                TableColumn ColumnNaturaleza = table.Columns[6];//3
                ColumnNaturaleza.Name = "NATURALEZA";

                TableColumn ColumnSabCompra = table.Columns[7];
                ColumnSabCompra.Name = "AGENTE COMPRA";

                TableColumn ColumnCompras = table.Columns[8];//4
                ColumnCompras.Name = "COMPRAS";

                TableColumn ColumnSabVenta = table.Columns[9];
                ColumnSabVenta.Name = "AGENTE VENTA";

                TableColumn ColumnVentas = table.Columns[10];//5
                ColumnVentas.Name = "VENTAS";

                TableColumn ColumnNroCertificado = table.Columns[11];
                ColumnNroCertificado.Name = "NRO. CERTIFICADO";

                ColumnCompras.DataRange.NumberFormat = "###,###,###,###,###,##0";
                ColumnVentas.DataRange.NumberFormat = "###,###,###,###,###,##0";

                table.HeaderRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                table.HeaderRowRange.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                CellRange RangoTotalFinal = worksheet["A5:L" + (dtDatosCartera.Rows.Count + 11) + ""];

                for (int i = 0; i < table.Columns.Count; i++)
                {
                    table.AutoFilter.Columns[i].HiddenButton = true;

                    switch (i)
                    {
                        case 0:
                            worksheet.Columns[i].Width = 260;
                            break;
                        case 1:
                            worksheet.Columns[i].Width = 260;
                            break;
                        case 2:
                            worksheet.Columns[i].Width = 260;
                            break;
                        case 3:
                            worksheet.Columns[i].Width = 260;
                            break;
                        case 4:
                            worksheet.Columns[i].Width = 1160;
                            break;
                        case 6:
                            worksheet.Columns[i].Width = 960;
                            break;
                        case 11:
                            worksheet.Columns[i].Width = 265;
                            break;
                        default:
                            worksheet.Columns[i].Width = 255;
                            break;
                    }
                }

                RangoTotalFinal = worksheet["A6:L" + (dtDatosCartera.Rows.Count + 5) + ""];
                List<int> subtotalColumnsList = new List<int>();
                subtotalColumnsList.Add(8);
                subtotalColumnsList.Add(10);

                worksheet.Subtotal(RangoTotalFinal, 6, subtotalColumnsList, 9, "TOTAL");


                var nroFechasDiferentes = dtDatosCartera.AsEnumerable().Select(a => new { fecha = a.Field<DateTime>("FEC_OPERACION"), naturaleza = a.Field<string>("DES_NATURALEZA") }).Distinct();
                var naturalezas = dtDatosCartera.AsEnumerable().Select(a => a.Field<string>("DES_NATURALEZA")).Distinct().ToList();
                var indiceFila = 5;

                for (int i = 0; i < naturalezas.Count(); i++)
                {
                    indiceFila += dtDatosCartera.AsEnumerable().Count(r => r.Field<String>("DES_NATURALEZA") == naturalezas[i]) + 1;

                    worksheet[$"I{ indiceFila }"].Font.Bold = true;//e
                    worksheet[$"K{ indiceFila }"].Font.Bold = true;//f
                }
                indiceFila++;

                //Cambiamos el nombre de la palabra "GRAND TOTAL" por Total General

                var celda = $"G{ indiceFila }";//d
                worksheet[celda].Value = "TOTAL GENERAL";
                worksheet[celda].Font.Color = Color.Black;

                celda = $"I{ indiceFila }";//e
                worksheet[celda].Font.Color = Color.Black;
                worksheet[celda].Font.Bold = true;

                celda = $"K{ indiceFila }";//f
                worksheet[celda].Font.Color = Color.Black;
                worksheet[celda].Font.Bold = true;

                worksheet.ActiveView.Orientation = PageOrientation.Landscape;
                worksheet.ActiveView.PaperKind = System.Drawing.Printing.PaperKind.A4;

            }
            workbook.Worksheets[0].Visible = false;
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[1];

            var model = new DocumentContentFromBytes
            {
                DocumentId = "ReporteActasTransferencia.xlsx",
                ContentAccessor = () => workbook.SaveDocument(DocumentFormat.Xlsx),
                IsEditable = true,
                FileName = "ReporteActasTransferencia.xlsx"
            };
            return View("~/Views/Tools/SpreadSheet.cshtml", model);
        }
    }
}