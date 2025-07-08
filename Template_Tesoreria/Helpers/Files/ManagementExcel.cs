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
//using Spire.Xls;


namespace Template_Tesoreria.Helpers.Files
{
    public class ManagementExcel
    {
        private string _path;
        private FileInfo _file;

        public ManagementExcel(string pathExcel) 
        {
            this._path = pathExcel;
            this._file = new FileInfo(this._path);
            ExcelPackage.License.SetNonCommercialOrganization("Grupo Sanborns");
        }

        public string cleanSheets(string sheet)
        {
            try
            {
                using(var package = new ExcelPackage(this._file))
                {
                    var sheetToClean = package.Workbook.Worksheets[sheet];
                    sheetToClean.Cells["A5:A15"].Delete(eShiftTypeDelete.EntireRow);
                    return "ELIMINADO";
                }
            }
            catch (Exception ex)
            {
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

                    foreach(var rows in data)
                    {
                        sheet.Cells[$"B{i}"].Value = rows.Cuenta ?? "";
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
                }
            }
            catch(Exception ex)
            {
                return $"Hubo un pequeño error: {ex.Message}";
            }
            return "Hubo un ligero error";
        }
    }
}
