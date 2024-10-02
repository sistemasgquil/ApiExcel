using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using Twilio.Jwt.AccessToken;
using OfficeOpenXml;

namespace ConsumirApiExternaExcel
{
    public partial class Form1 : Form
    {
        string rutaArchivo, urlApiProveedor;
         
        private static readonly HttpClient client = new HttpClient();


        public System.Net.Http.Headers.HttpRequestHeaders DefaultRequestHeaders { get; }

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();// eND

        }
        private static async Task PostLlamadoExcelAsync(string fileName, string filePath)
        {
            
            // Crear una instancia de HttpClient
            using (HttpClient client = new HttpClient())
            {
                // Establecer el header de la suscripción   TOKEN
                client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "9f946f01490f41e3b85968572d72c1fa");

                // URL de la API con el nombre del archivo
                string url = $"https://api-us.vusion.io/vcloud/v1/stores/retailpoint_ec.lab/items/files/{fileName}";

                // Leer el archivo que se quiere enviar
                byte[] fileBytes =  File.ReadAllBytes(filePath);

                // Crear el contenido para la solicitud (octet-stream)
                ByteArrayContent content = new ByteArrayContent(fileBytes);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                //Ocp-Apim-Subscription-Key        9f946f01490f41e3b85968572d72c1fa

                // Enviar la solicitud POST
                HttpResponseMessage response = await client.PostAsync(url, content);

                // Mostrar el resultado
                if (response.IsSuccessStatusCode)
                {
     
                    MessageBox.Show("Archivo enviado exitosamente.","Mensaje al Usuario");
                }
                else
                {
                    MessageBox.Show($"Error al enviar el archivo: { response.StatusCode}", "Mensaje al Usuario");
                    
                }
            }




            //await Task.Delay(3000); // Esperar 3 segundos
        }

            private async void btnEjecutar_Click(object sender, EventArgs e)
            {
            try
            {
                // Crear el dialogo de selección de archivo
                OpenFileDialog openFileDialog = new OpenFileDialog();
                // Filtrar solo archivos Excel (XLSX)
                openFileDialog.Filter = "Excel Files|*.csv;";
                openFileDialog.Title = "Selecciona un archivo CSV";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Obtener la ruta completa del archivo seleccionado
                    string rutaArchivo = openFileDialog.FileName;
                    string nombreArchivo= Path.GetFileName(rutaArchivo);
                    txtRutaArchivo.Text = rutaArchivo;

                    MessageBox.Show("Archivo Seleccionado: " + rutaArchivo);
                    if (txtRutaProveedor.Text.Length < 1 && txtRutaArchivo.Text.Length < 1) { return; }
                    urlApiProveedor = txtRutaProveedor.Text.Trim();
                    await PostLlamadoExcelAsync(nombreArchivo, rutaArchivo);
                   // // await PostLlamadoExcelAsync("BORX26030924-01.csv", rutaArchivo);
                }
            }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message); return;
                }
            }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            btnEjecutar.Focus();
        }

        #region"Leer data interna de la apis No usar para este caso, x q solo consume servicio del proveedor"



        public async Task LeerExcelYEnviarDatosAPI(string rutaArchivo)
        {
            // Configurar el contexto de licencia para EPPlus
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Abrir el archivo de Excel
            using (var package = new ExcelPackage(new FileInfo(rutaArchivo)))
            {
                // Obtener la primera hoja de cálculo
                var worksheet = package.Workbook.Worksheets[0];

                // Obtener el número total de filas con datos
                int rowCount = worksheet.Dimension.Rows;

                // Recorrer las filas desde la fila 2 (asumiendo que la fila 1 contiene los encabezados)
                for (int row = 2; row <= rowCount; row++)
                {
                    // Leer los valores de cada columna
                    string descripcion = worksheet.Cells[row, 1].Text;
                    string referencia = worksheet.Cells[row, 2].Text;
                    string medida = worksheet.Cells[row, 3].Text;
                    string valor = worksheet.Cells[row, 4].Text;

                    // Mostrar los datos leídos (opcional)
                    Console.WriteLine($"Fila {row} - Descripción: {descripcion}, Referencia: {referencia}, Medida: {medida}, Valor: {valor}");

                    // Enviar los datos a la API
                    await EnviarDatosAPI(descripcion, referencia, medida, valor);
                }
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        // Método para enviar los datos leídos a la API
        private async Task EnviarDatosAPI(string descripcion, string referencia, string medida, string valor)
        {
            var jsonData = new
            {
                Descripcion = descripcion,
                Referencia = referencia,
                Medida = medida,
                Valor = valor
            };

            var jsonContent = new StringContent(
                Newtonsoft.Json.JsonConvert.SerializeObject(jsonData),
                Encoding.UTF8,
                "application/json"
            );

            var response = await client.PostAsync("https://tuapi.com/endpoint", jsonContent);

            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine("Datos enviados correctamente.");
            }
            else
            {
                Console.WriteLine($"Error al enviar los datos: {response.StatusCode}");
            }
        }
        #endregion
    }

}

