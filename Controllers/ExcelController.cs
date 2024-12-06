using api_excel_new.Services;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace api_excel_new.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly ApiService _apiService;

        public ExcelController()
        {
            _apiService = new ApiService();
        }

        [HttpPost("ReadExcelFromBase64")]
        public IActionResult ReadExcelFromBase64([FromBody] ExcelRequestModel model)
        {
            if (model == null || string.IsNullOrEmpty(model.Base64))
                return BadRequest("No se proporcionó un archivo en Base64.");

            try
            {
                var fileBytes = Convert.FromBase64String(model.Base64);
                var result = new Dictionary<string, object>();
                string[] datosNames = ["datos_empresa", "datos_cliente", "formato_venta"];

                using (var stream = new MemoryStream(fileBytes))
                {
                    using (var workbook = new XLWorkbook(stream))
                    {
                        var sheetVendedor = workbook.Worksheet("Cotizador vendedor");
                        var sheetVemdo = workbook.Worksheet("Datos para Vemdo");
                        if (sheetVendedor == null && sheetVemdo == null)
                        {
                            return BadRequest("Las hojas requeridas no existen en el archivo Excel.");
                        }
                        var tables = new Dictionary<string, (string startCell, string endCell)>
                    {
                        { "fecha", ("E4", "E4") },
                        { "datos_empresa", ("B8", "C10") },
                        { "datos_cliente", ("B14", "C16") },
                        { "formato_venta", ("B29", "C33") },
                        { "comentarios_vendedor", ("B20", "C25") },
                        { "table_data", ("E8", "I25") },
                        { "total", ("H27", "I34") },
                        { "valores_venta_comision", ("C2", "C3") },
                        { "moneda", ("H4", "H4") }
                    };

                        foreach (var table in tables)
                        {
                            (string startCell, string endCell) = table.Value;
                            //string startCell = "A1";
                            //string endCell = "J34";
                            if (table.Key == "total")
                            {
                                result[table.Key] = _apiService.TransformTotalTable(sheetVendedor);
                            }
                            else if (table.Key == "valores_venta_comision")
                            {
                                result[table.Key] = _apiService.ExtractValoresVentaComision(sheetVemdo);
                            }
                            else if (table.Key == "moneda")
                            {
                                result[table.Key] = _apiService.extract_moneda(sheetVendedor);
                            }
                            else
                            {
                                //result[table.Key] = _apiService.ExtractTable(sheetVendedor, startCell, endCell);
                                var data = _apiService.ExtractTable(sheetVendedor, startCell, endCell);
                                //return Ok(data);
                                if (table.Key == "fecha")
                                {
                                    //string fecha = data[3][4]?.ToString();
                                    result[table.Key] = _apiService.transform_fecha(data);

                                }

                                else if (datosNames.Contains(table.Key))
                                {
                                    //Debug.WriteLine("Entrando");
                                    //var stringData = data.Select(row => row.Select(item => item?.ToString()).ToList()).ToList();
                                    result[table.Key] = _apiService.TransformLeftRightTable(data, table.Key);
                                    //return Ok(result);
                                }
                                else if (table.Key == "comentarios_vendedor")
                                {
                                    result[table.Key] = _apiService.TransformComentariosVendedor(data);
                                    //return Ok(result);
                                }
                                else if (table.Key == "table_data")
                                {
                                    result[table.Key] = _apiService.TransformTableData(data);
                                }
                            }

                        }


                        var keysToClean = new List<string> { "datos_empresa", "datos_cliente", "formato_venta", "table_data" };
                        var keysToRemoveInObjects = new HashSet<string> { "Datos Empresa", "Datos Cliente", "Formato de la venta", "null" };

                        foreach (var key in keysToClean)
                        {
                            if (result.ContainsKey(key))
                            {
                                var value = result[key];

                                if (value is Dictionary<string, object> dictValue)
                                {
                                    result[key] = dictValue
                                        .Where(pair => !keysToRemoveInObjects.Contains(pair.Key) || (pair.Value != null && pair.Value.ToString() != string.Empty))
                                        .ToDictionary(pair => pair.Key, pair => pair.Value);
                                }
                                else if (value is List<Dictionary<string, object>> listValue)
                                {
                                    result[key] = listValue
                                        .Select(item => item
                                            .Where(pair => pair.Key != null)
                                            .ToDictionary(pair => pair.Key, pair => pair.Value))
                                        .ToList();
                                }
                            }
                        }
                    }
                }

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error procesando el archivo: {ex.Message}");
            }
        }
    }

}
public class ExcelRequestModel
{
    public string Base64 { get; set; }
    //public string StartCell { get; set; }
    //public string EndCell { get; set; }
}