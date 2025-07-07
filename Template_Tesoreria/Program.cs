using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using HTML = HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Net;
using Template_Tesoreria.Helpers.DataAccess;
using Template_Tesoreria.Models;
using Template_Tesoreria.Helpers.Files;

namespace Template_Tesoreria
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var dtService = new DataService();
            var cnn = new ConnectionDb();
            string opc = "", opc2 = "", nombreBanco = "", rutaCarpeta = "", urlArchivoDescaga = "", pathDestino = "";
            int numero;

            try
            {
                while (true)
                {
                    Console.Write("\nSelecciona la Compañia de Activo que deseas generar \n 1 - Inbursa \n 2 - HSBC \n 3 - Bancomer \n 4 - Scotiabank \n 5 - Citi \n 6 - Santander \n 7 - Banorte\n");
                    Console.Write("Opcion: ");
                    opc = Console.ReadLine().Trim();

                    if (int.TryParse(opc, out numero))
                    {
                        numero = Convert.ToInt32(opc);
                        if (numero == 1 || numero == 2 || numero == 3 || numero == 4 || numero == 5 || numero == 6 || numero == 7)
                        {
                            if (numero == 1) nombreBanco = "Inbursa";
                            if (numero == 2) nombreBanco = "HSBC";
                            if (numero == 3) nombreBanco = "Bancomer";
                            if (numero == 4) nombreBanco = "Scotiabank";
                            if (numero == 5) nombreBanco = "Citibanamex";
                            if (numero == 6) nombreBanco = "Santander";
                            if (numero == 7) nombreBanco = "Banorte";

                            Console.Write("\nEsta seguro de querer trabajar con: " + nombreBanco + "\n");
                            Console.Write(" 1 - SI\n 2 - NO \n");
                            Console.Write("Opcion: ");
                            opc2 = Console.ReadLine().Trim();

                            if (opc2.Equals("1")) break;
                            //break;
                        }
                        else
                        {
                            Console.Write("No existe la Opción " + opc + " en la lista\n\n");
                        }
                    }
                    else
                    {
                        Console.Write("El valor que ingreso no es Númerico!!!\n\n");
                    }
                }

                Console.Write("\nDescargando Template de Oracle\n\n");                    

                WebClient client1 = new WebClient();
                string htmlCode = client1.DownloadString("https://docs.oracle.com/en/cloud/saas/financials/25b/oefbf/cashmanagementbankstatementdataimport-3168.html#cashmanagementbankstatementdataimport-3168");
                string[] lines = htmlCode.Split('\n');

                HTML.HtmlDocument htmlDocument = new HTML.HtmlDocument();
                htmlDocument.LoadHtml(lines[58].ToString().Trim());

                var linkNodes = htmlDocument.DocumentNode.SelectNodes("//a[@href]");

                if (linkNodes != null)
                {
                    foreach (var linkNode in linkNodes)
                    {
                        urlArchivoDescaga = linkNode.GetAttributeValue("href", string.Empty);
                        //Console.WriteLine($"Enlace: {hrefValue}");
                    }
                }

                rutaCarpeta = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Downloads\\Templates";

                //Si no existe la Carpeta la creamos
                if (!Directory.Exists(rutaCarpeta)) Directory.CreateDirectory(rutaCarpeta);

                //Definimos la ruta donde guardaremos el archivo
                //http://www.oracle.com/webfolder/technetwork/docs/fbdi-25b/fbdi/xlsm/CashManagementBankStatementImportTemplate.xlsm                
                pathDestino = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Downloads\\Templates\\CashManagementBankStatementImportTemplate_" + nombreBanco + ".xlsm";

                WebClient myWebClient = new WebClient();
                myWebClient.DownloadFile(urlArchivoDescaga, pathDestino);

                //COLOCAR EL LLENADO DEL EXCEL
                var data = dtService.GetDataList<Tbl_Tesoreria_Ext_Bancario>(cnn.DbTesoreria1019(), "pa_Tesoreria_SelDatos", null);
                var mngmntExcel = new ManagementExcel(pathDestino);
                var fillData = mngmntExcel.getTemplate(data);


                Console.Write("\nTemplate de Oracle Descargado con Exito\n\n");


                //Proceso para Leer Formato de Banco
                //UploadFile("");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void UploadFile(string rutaFormato)
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(System.Configuration.ConfigurationManager.AppSettings.Get("FtpRuta"));
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.Credentials = new NetworkCredential(System.Configuration.ConfigurationManager.AppSettings.Get("FtpUser"), System.Configuration.ConfigurationManager.AppSettings.Get("FtpPass"));

                byte[] fileContents = File.ReadAllBytes(rutaFormato);
                request.ContentLength = fileContents.Length;

                using (Stream requestStream = request.GetRequestStream())
                {
                    requestStream.Write(fileContents, 0, fileContents.Length);
                }

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                {
                    Console.WriteLine($"Upload File Complete, status {response.StatusDescription}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
    }
}
