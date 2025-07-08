//using Microsoft.Office.Interop.ExcKel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Template_Tesoreria.Models;
using OfficeOpenXml;
using System.IO;
using System.ComponentModel.DataAnnotations;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
//using Spire.Xls;


namespace Template_Tesoreria.Helpers.Files
{
    public class ManagementExcel
    {
        private string _path;
        private FileInfo _file;
        private Log _log;

        public ManagementExcel(string pathExcel) 
        {
            this._path = pathExcel;
            this._file = new FileInfo(this._path);
            ExcelPackage.License.SetNonCommercialOrganization("Grupo Sanborns");
            this._log = new Log();
        }

        public string cleanSheets(string sheet)
        {
            this._log.writeLog($"LIMPIEZA DE LA HOJA {sheet}");
            try
            {
                using(var package = new ExcelPackage(this._file))
                {
                    var sheetToClean = package.Workbook.Worksheets[sheet];
                    sheetToClean.DeleteRow(5, 15);
                    package.Save();
                    this._log.writeLog($"LIMPIEZA TERMINADA, TODO CORRECTO");
                    return "ELIMINADO";
                }
            }
            catch (Exception ex)
            {
                this._log.writeLog($"HUBO UN LIGERO ERROR AL QUERER LIMPIAR LA HOJA {sheet}\n\t\tERROR: {ex.Message}");
                return ex.Message;
            }
        }

        public string getTemplate(List<Tbl_Tesoreria_Ext_Bancario> data)
        {
            try
            {
                using(var package = new ExcelPackage(this._file))
                {
                    var sheet = package.Workbook.Worksheets["Statement Lines"];
                    int i = 5;
                    
                    this._log.writeLog($"COMIENZO CON CICLO PARA LA INSERCIÓN DE DATOS.\n\t\tSE INSERTARAN {data.Count} REGISTROS");

                    foreach (var rows in data)
                    {
                        sheet.Cells[$"B{i}"].Value = rows.Cuenta.Replace("-PESOS", "") ?? "";
                        sheet.Cells[$"D{i}"].Value = rows.Concepto ?? "";
                        sheet.Cells[$"H{i}"].Value = rows.Fecha ?? "";
                        sheet.Cells[$"L{i}"].Value = rows.Referencia ?? "";
                        sheet.Cells[$"S{i}"].Value = rows.RFC_Ordenante ?? "";
                        sheet.Cells[$"T{i}"].Value = rows.Ordenante ?? "";
                        sheet.Cells[$"W{i}"].Value = rows.Movimiento ?? "";
                        sheet.Cells[$"X{i}"].Value = rows.Referencia_Leyenda ?? "";
                        sheet.Cells[$"BN{i}"].Value = rows.Referencia_Ext ?? "";
                        sheet.Cells[$"BP{i}"].Value = rows.Referencia_Numerica ?? "";
                        i++;
                    }
                    package.Save();
                    this._log.writeLog($"SE INSERTARON LOS REGISTROS CORRECTAMENTE");
                    return "CORRECTO";
                }
            }
            catch(Exception ex)
            {
                this._log.writeLog($"HUBO UN LIGERO ERROR AL INSERTAR LOS DATOS\n\t\tERROR: {ex.Message}");
                return $"Hubo un pequeño error: {ex.Message}";
            }
        }

        public void closeDocument()
        {
            Excel.Application excelApp = null;

            var index = this._path.LastIndexOf(@"\\");
            var file = "";

            if (index != -1)
                file = this._path.Substring(index + 1);

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                foreach(Excel.Workbook wb in excelApp.Workbooks)
                {
                    if(wb.FullName.EndsWith(file))
                    {
                        wb.Close(true);
                        break;
                    }
                }

                if (excelApp.Workbooks.Count == 0)
                    excelApp.Quit();
            }
            catch(Exception ex)
            {
                this._log.writeLog($"HUBO UN PEQUEÑO ERROR AL QUERER CERRAR EL DOCUMENTO DE EXCEL\n\t\tERROR: {ex.Message}");
            }
        }
    }
}
