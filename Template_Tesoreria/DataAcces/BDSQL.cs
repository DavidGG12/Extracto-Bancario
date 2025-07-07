using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.DataAcces
{
    public class BDSQL
    {
        public DataSet CmdDeDataSet(string sentenciaSQL, string DataBase)
        {
            BD oBD = new BD();
            DataSet ds = new DataSet();
            oBD.Conectar(DataBase);
            oBD.CrearComando(sentenciaSQL);
            ds = oBD.DataReader_DataSet();
            oBD.Desconectar();
            return ds;
        }

        public string CmdDeDataResultado(string sentenciaSQL, string DataBase, string Campo)
        {
            string Resultado = "";
            BD oBD = new BD();
            oBD.Conectar(DataBase);
            oBD.CrearComando(sentenciaSQL);
            try
            {
                DbDataReader reader = oBD.EjecutarConsulta();
                while (reader.Read())
                {
                    Resultado = reader[Campo].ToString();
                }
                reader.Close();
            }
            catch { }
            oBD.Desconectar();
            return Resultado;
        }

        public int CmdDeEjecucionEscalar(string sentenciaSQL, string DataBase)
        {
            int Valorescalar;
            BD oBD = new BD();
            oBD.Conectar(DataBase);
            oBD.CrearComando(sentenciaSQL);
            Valorescalar = oBD.EjecutarEscalar();
            oBD.Desconectar();
            return Valorescalar;
        }

        public void CmdDeEjecucion(string sentenciaSQL, string DataBase)
        {
            BD oBD = new BD();
            oBD.Conectar(DataBase);
            oBD.CrearComando(sentenciaSQL);
            oBD.EjecutarComando();
            oBD.Desconectar();
        }
    }
}
