using System;
using System.ComponentModel;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;


namespace ConsumirApiExternaExcel
{

    public class ExcelUploader
    {
    public string respuestaApi;

        #region"Metodos para consumir la API de Excel"
        // Método para enviar el archivo Excel a la API del proveedor
        public async Task EnviarArchivoExcelAPI(string rutaArchivo, string urlApiProveedor)
        {
            // Crear el cliente HTTP
            using (var client = new HttpClient())
            {
                // Crear el contenido multipart/form-data
                using (var content = new MultipartFormDataContent())
                {
                    // Abrir el archivo de Excel
                    var archivoStream = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read);
                    var archivoContent = new StreamContent(archivoStream);

                    // Configurar el tipo MIME del archivo
                    archivoContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                    // Agregar el archivo al contenido de la solicitud
                    content.Add(archivoContent, "file", Path.GetFileName(rutaArchivo));

                    // Hacer la solicitud POST a la API del proveedor
                    var response = await client.PostAsync(urlApiProveedor, content);

                    // Validar si la solicitud fue exitosa
                    if (response.IsSuccessStatusCode)
                    {
                        // Console.WriteLine("Archivo enviado correctamente.");
                        respuestaApi = "Archivo enviado correctamente.";
                    }
                    else
                    {
                        // Console.WriteLine($"Error al enviar el archivo: {response.StatusCode}");
                            respuestaApi = $"Error al enviar el archivo: {response.StatusCode}";
                    }
                }
            }
        }

        #endregion

       

    }
}
