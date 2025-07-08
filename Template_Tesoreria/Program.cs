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
using Template_Tesoreria.Helpers.ProcessExe;
using System.Diagnostics;
using System.Threading;

namespace Template_Tesoreria
{
    internal class Program
    {
        static void Spinner(string message, CancellationToken token)
        {
            char[] secuence = { '|', '/', '-', '\\' };
            int pos = 0;

            while(!token.IsCancellationRequested)
            {
                Console.Write($"\r{message} {secuence[pos]}");
                pos = (pos + 1) % secuence.Length;
                Thread.Sleep(100);
            }
            Console.Write($"\r{new string(' ', Console.WindowWidth)}");
            Console.Write("\rTerminado\n");
        }

        static void Main(string[] args)
        {
            var dtService = new DataService();
            var cnn = new ConnectionDb();
            var cts = new CancellationTokenSource();
            var log = new Log();
            var process = new ProcessPython(@"\\10.115.0.14\Finanzas\Tesoreria\EXE\getTablesAccounts.exe");
            string opc = "", opc2 = "", nombreBanco = "", rutaCarpeta = "", urlArchivoDescaga = "", pathDestino = "";
            int numero;

            try
            {
                log.writeLog("COMENZANDO PROCESO");

                while (true)
                {
                    log.writeLog("IMPRESIÓN DEL MENÚ");

                    Console.Write("\nSelecciona la Compañia de Activo que deseas generar \n 1 - Inbursa \n 2 - HSBC \n 3 - Bancomer \n 4 - Scotiabank \n 5 - Citi \n 6 - Santander \n 7 - Banorte\n");
                    Console.Write("Opcion: ");
                    opc = Console.ReadLine().Trim();

                    log.writeLog($"SE ESCOGIÓ LA OPCIÓN: {opc}");

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

                            log.writeLog($"LA OPCIÓN {opc} CORRESPONDE AL BANCO: {nombreBanco}");

                            Console.Write($"\n¿Está seguro de querer trabajar con {nombreBanco}? [S/N]: ");
                            opc2 = Console.ReadLine().Trim();


                            if (opc2.Equals("s", StringComparison.OrdinalIgnoreCase))
                            {
                                log.writeLog($"SE CONFIRMA EL USO DEL BANCO {nombreBanco}");
                                break;
                            }
                        }
                        else
                        {
                            Console.Write("No existe la Opción " + opc + " en la lista\n\n");
                            log.writeLog($"SE INGRESÓ LA OPCIÓN: {opc} Y NO SE ENCUENTRA DENTRO DE LA LISTA");
                        }
                    }
                    else
                    {
                        Console.Write("El valor que ingreso no es Númerico!!!\n\n");
                    }
                }

                Console.Write("\nDescargando Template de Oracle\n\n");
                log.writeLog($"SE EMPIENZA EL LLENADO DEL TEMPLATE DE ORACLE");

                var result = "";

                Task.Run(() =>
                    {
                        result = process.ExecuteProcess();
                        cts.Cancel();
                    }
                );

                Spinner("Procesando...", cts.Token);

                if (!string.Equals(result.TrimEnd().TrimStart(), "DATOS INSERTADOS CORRECTAMENTE"))
                {
                    Console.WriteLine(result);
                    log.writeLog($"HUBO UN LIGERO ERROR AL QUERER INSERTAR LOS DATOS\n\tERROR: {result}");
                    return;
                }

                Console.WriteLine("\nDatos descargados.\n\n");
                
                log.writeLog($"LOS DATOS SE HAN DESCARGADO CORRECTAMENTE");
                log.writeLog($"COMIENZA LA DESCARGA DEL TEMPLATE");

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

                log.writeLog($"SE OBTUVO LA INFORMACIÓN PARA PODER DESCARGAR CORRECTAMENTE EL TEMPLATE");

                rutaCarpeta = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Downloads\\Templates";

                //Si no existe la Carpeta la creamos
                if (!Directory.Exists(rutaCarpeta)) Directory.CreateDirectory(rutaCarpeta);


                //Definimos la ruta donde guardaremos el archivo
                //http://www.oracle.com/webfolder/technetwork/docs/fbdi-25b/fbdi/xlsm/CashManagementBankStatementImportTemplate.xlsm                
                pathDestino = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Downloads\\Templates\\CashManagementBankStatementImportTemplate_" + nombreBanco + ".xlsm";
                var mngmntExcel = new ManagementExcel(pathDestino);

                log.writeLog($"EL TEMPLATE SE INSERTARÁ EN LA SIGUIENTE RUTA: {pathDestino}");
                
                WebClient myWebClient = new WebClient();

                Task.Run(() =>
                {
                    mngmntExcel.closeDocument();
                    myWebClient.DownloadFile(urlArchivoDescaga, pathDestino);
                    cts.Cancel();
                }
                );

                Spinner("Procesando...", cts.Token);

                log.writeLog($"SE DESCARGA EL TEMPLATE");
                log.writeLog($"EMPIEZA LA INSERCIÓN DE LOS DATOS EN EL TEMPLATE");

                //Empezamos con la recolección de datos y el llenado de la información
                var data = dtService.GetDataList<Tbl_Tesoreria_Ext_Bancario>(cnn.DbTesoreria1019(), "pa_Tesoreria_SelDatos", null);


                //Limpiamos el template para trabajar con él
                log.writeLog($"LIMPIAMOS EL TEMPLATE PARA PODER INSERTAR LOS DATOS");
                var errorList = new List<SheetError>();
                errorList.Add(new SheetError() { Sheet = "Statement Headers", Message = mngmntExcel.cleanSheets("Statement Headers") });
                errorList.Add(new SheetError() { Sheet = "Statement Balances", Message = mngmntExcel.cleanSheets("Statement Balances") });
                errorList.Add(new SheetError() { Sheet = "Statement Balance Availability", Message = mngmntExcel.cleanSheets("Statement Balance Availability") });
                errorList.Add(new SheetError() { Sheet = "Statement Lines", Message = mngmntExcel.cleanSheets("Statement Lines") });
                errorList.Add(new SheetError() { Sheet = "Statement Line Avilability", Message = mngmntExcel.cleanSheets("Statement Line Availability") });
                errorList.Add(new SheetError() { Sheet = "Statement Statement Line Charges", Message = mngmntExcel.cleanSheets("Statement Line Charges") });

                var error = errorList.Find(x => !x.Message.Contains("ELIMINADO"));
                if(error != null)
                {
                    Console.WriteLine($"Hubo un ligero error al querer limpiar los datos de la hoja {error.Sheet}.\nError: {error.Message}");
                    return;
                }

                log.writeLog($"TERMINO DE LIMPIEZA, SE PROSIGUE CON LA INSERCIÓN DE DATOS");

                //Insertamos los datos que se encuentran en la base de datos
                var fillData = mngmntExcel.getTemplate(data);

                Console.Write("\nTemplate de Oracle Descargado con Exito\n\n");

                Process.Start(pathDestino);
                log.writeLog($"ABRIENDO ARCHIVO\n\t\t**PROCESO TERMINADO**");
                log.writeLog($"**********************************************************************");

                //Proceso para Leer Formato de Banco
                //UploadFile("");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                log.writeLog($"ALGO OCURRIÓ DURANTE EL PROCESO PRINCIPAL {ex.Message}");
                log.writeLog($"**********************************************************************");
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
