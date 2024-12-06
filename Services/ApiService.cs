using ClosedXML.Excel;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;

namespace api_excel_new.Services
{
    public class ApiService
    {
        public List<List<object>> ExtractTable(IXLWorksheet sheet, string startCell, string endCell)
        {
            Debug.WriteLine("StartCell: " + startCell + " EndCell: " + endCell);
            var data = new List<List<object>>();
            var range = sheet.Range(startCell, endCell);

            foreach (var row in range.Rows())
            {
                var rowData = new List<object>();
                foreach (var cell in row.Cells())
                {
                    rowData.Add(cell.GetValue<string>());
                }
                data.Add(rowData);
            }
            string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            Debug.WriteLine(json);
            return data;
        }

        public Dictionary<string, object> TransformTotalTable(IXLWorksheet sheet)
        {
            Debug.WriteLine("TOTAL", sheet.Cell("I33").Value);
            Dictionary<string, object> data = new Dictionary<string, object>
            {
                { "SUB-TOTAL", sheet.Cell("I27").Value.ToString() },
                { "IVA (19%)", sheet.Cell("I28").Value.ToString() },
                { "Total", sheet.Cell("I29").Value.ToString() },
                { "Descuento empresa (%)", sheet.Cell("H31").Value.ToString() },
                { "Descuento monto", sheet.Cell("I31").Value.ToString() },
                { "TOTAL", sheet.Cell("I33").Value.ToString() }
            };
            string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            Debug.WriteLine("Total " + json);
            return data;
        }

        public Dictionary<string, object> ExtractValoresVentaComision(IXLWorksheet sheet)
        {
            return new Dictionary<string, object>
            {
                { "Valor venta (con IVA)", sheet.Cell("B1").Value.ToString() },
                { "Valor comisión (con IVA)", sheet.Cell("B2").Value.ToString() }
            };
        }
        public Dictionary<string, string> extract_moneda(IXLWorksheet sheet)
        {
            var dictMoneda = new Dictionary<string, string>
            {
                { "$ Peso Chileno", "CLP" },
                { "UF", "CLF" }
            };
            var monedaValue = sheet.Cell("H4").GetValue<string>();
            if (string.IsNullOrWhiteSpace(monedaValue))
            {
                throw new Exception("La moneda no se pudo extraer: el contenido esperado no está en la celda H4.");
            }
            monedaValue = dictMoneda[monedaValue];
            Debug.WriteLine("Moneda value " + monedaValue);
            return new Dictionary<string, string> { { "Moneda", monedaValue } };
        }
        public string transform_fecha(object data)
        {
            string dataAux = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            Debug.WriteLine(dataAux);
            if (data == null)
            {
                throw new Exception("Error al procesar la fecha: la data está vacía o es inválida.");
            }
            List<List<object>> newData = JsonSerializer.Deserialize<List<List<object>>>(dataAux);
            object dateObject = newData[0][0];
            string dateString = dateObject.ToString();
            Debug.WriteLine(dateObject);
            try
            {
                if (dateString == string.Empty)
                {
                    throw new Exception("Por favor coloque una fecha");
                }
                DateTime dateValue;
                if (data is DateTime datetime)
                {
                    return datetime.ToString("yyyy-MM-dd");
                }

                var tryFormats = new List<string>
                {
                    "yyyy-MM-dd",
                    "dd-MM-yyyy",
                    "dd/MM/yyyy",
                    "yyyy/MM/dd",
                    "dd-MM-yy",
                    "dd/MM/yy",
                    "yyyy.MM.dd",
                    "dd.MM.yyyy"
                };
                foreach (var format in tryFormats)
                {
                    if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue))
                    {
                        return dateValue.ToString("yyyy-MM-dd");
                    }
                }
                if (DateTime.TryParse(dateString, out dateValue))
                {
                    return dateValue.ToString("yyyy-MM-dd");
                }
                throw new Exception($"El formato de fecha '{data}' no es válido.");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al procesar la fecha: {ex.Message}");
            }
        }
        public Dictionary<string, string> TransformLeftRightTable(List<List<object>> data, string name)
        {
            //if (data[0][0].ToString())
            Debug.WriteLine("Name: " + name);
            Dictionary<string, string> dataAux = data
                .Where(row => row.Count > 0 && row[0] != null && !string.IsNullOrWhiteSpace(row[0].ToString()))
                .ToDictionary(
                    row => row[0].ToString(),
                    row => row.Skip(1)
                        .FirstOrDefault(value => value != null && !string.IsNullOrWhiteSpace(value?.ToString()))?.ToString()
                );

            if (name == "datos_empresa"
                || name == "datos_cliente")
            {
                foreach (KeyValuePair<string, string> aux in dataAux)
                {
                    if (string.IsNullOrWhiteSpace(aux.Value))
                    {
                        Debug.WriteLine("Valor : " + aux.Key + " " + aux.Value);
                        throw new Exception($"El campo: {aux.Key} No debe estar vacio");
                    }
                }
            }
            return dataAux;
        }



        public Dictionary<string, string> TransformComentariosVendedor(List<List<object>> data)
        {
            // Concatenar todos los valores no nulos en la lista de datos
            var comments = string.Join(" ", data
                .SelectMany(row => row) // Aplanar la lista de listas
                .Where(cell => cell != null) // Filtrar celdas no nulas
                .Select(cell => cell.ToString()) // Convertir cada celda a string
                .ToArray()); // Convertir la colección filtrada a un array para usar Join

            // Si hay comentarios, devolverlos, de lo contrario devolver null
            return new Dictionary<string, string>
            {
                { "Comentarios vendedor", string.IsNullOrWhiteSpace(comments) ? null : comments }
            };
        }
        public List<Dictionary<string, object>> TransformTableData(List<List<object>> data)
        {
            var headers = data[0];

            var result = data.Skip(1)
                             .Select(row => headers
                                 .Zip(row, (header, value) => new { header, value })
                                 .ToDictionary(x => x.header.ToString(), x => x.value))
                             .ToList();
            foreach (Dictionary<string, object> aux in result)
            {
                //foreach(var kvp in aux)
                //{
                //    Debug.WriteLine(kvp.Key + " " + kvp.Value);
                //    if (string.IsNullOrWhiteSpace(kvp.Value.ToString()))
                //    {
                //        throw new Exception($"el campo: {kvp.Key} no puede estar vacio");
                //    }
                //}
                string aux1 = aux["Unidades"].ToString();
                string aux2 = aux["Unidades"].ToString();
                if (!string.IsNullOrEmpty(aux["Unidades"].ToString()))
                {
                    if (!float.TryParse(aux1, out float value)
                    && !int.TryParse(aux2, out int value2))
                    {
                        Debug.WriteLine("Valor : " + aux["Unidades"].ToString());
                        throw new Exception($"Las unidades no pueden contener letras");
                    }
                }
            }
            return result;
        }
    }
}
