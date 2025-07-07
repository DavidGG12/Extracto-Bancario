using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.DataAcces
{
    public class BD
    {
        private DbConnection conexion = null;
        private DbCommand comando = null;
        private DbTransaction transaccion = null;
        private string cadenaConexion;

        private static DbProviderFactory factory = null;

        /// <summary>
        /// Crea una instancia del acceso a la base de datos.
        /// </summary>
        public BD()
        {
            Configurar();
        }

        #region Configuración de Acceso
        /// <summary>
        /// Configura el acceso a la base de datos para su utilización.
        /// </summary>
        /// <exception cref="BaseDatosException">Si existe un error al cargar la configuración.</exception>
        private void Configurar()
        {
            try
            {
                string proveedor = System.Configuration.ConfigurationManager.AppSettings.Get("ADONET");
                DbProviderFactories.ReferenceEquals(proveedor, System.Data.SqlClient.SqlClientFactory.Instance);
                cadenaConexion = System.Configuration.ConfigurationManager.AppSettings.Get("Conexion");
                BD.factory = DbProviderFactories.GetFactory(proveedor);
            }
            catch (Exception ex)
            {
                throw new BDException("Error al cargar la configuración del acceso a datos.", ex);
            }
        }
        #endregion

        #region Conectar y Desconectar
        /// <summary>
        /// Se concecta con la base de datos.
        /// </summary>
        /// <exception cref="BaseDatosException">Si existe un error al conectarse.</exception>
        public void Conectar(string DataBase)
        {
            if (conexion != null && !conexion.State.Equals(ConnectionState.Closed))
            {
                throw new BDException("La conexión ya se encuentra abierta.");
            }
            try
            {
                if (conexion == null)
                {
                    conexion = factory.CreateConnection();
                    conexion.ConnectionString = cadenaConexion + DataBase;
                }
                conexion.Open();
            }
            catch (DataException ex)
            {
                throw new BDException("Error al conectarse a la base de datos.", ex);
            }
        }

        /// <summary>
        /// Permite desconectarse de la base de datos.
        /// </summary>
        public void Desconectar()
        {
            if (conexion != null && conexion.State.Equals(ConnectionState.Open))
            {
                conexion.Close();
            }
        }
        #endregion

        #region Crear Comando
        /// <summary>
        /// Crea un comando en base a una sentencia SQL.
        /// Ejemplo:
        /// <code>SELECT * FROM Tabla WHERE campo1=@campo1, campo2=@campo2</code>
        /// Guarda el comando para el seteo de parámetros y la posterior ejecución.
        /// </summary>
        /// <param name="sentenciaSQL">La sentencia SQL con el formato: SENTENCIA [param = @param,]</param>
        public void CrearComando(string sentenciaSQL)
        {
            comando = factory.CreateCommand();
            comando.Connection = conexion;
            comando.CommandType = CommandType.Text;
            comando.CommandText = sentenciaSQL;
            comando.CommandTimeout = conexion.ConnectionTimeout;
            if (transaccion != null)
            {
                comando.Transaction = transaccion;
            }
        }
        #endregion

        #region Asignar Parametros
        /// <summary>
        /// Asigna un parámetro al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="separador">El separador que será agregado al valor del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        private void AsignarParametro(string nombre, string separador, string valor)
        {
            int indice = comando.CommandText.IndexOf(nombre);
            string prefijo = comando.CommandText.Substring(0, indice);
            string sufijo = comando.CommandText.Substring(indice + nombre.Length);
            comando.CommandText = prefijo + separador + valor + separador + sufijo;
        }

        /// <summary>
        /// Asigna un parámetro de tipo cadena al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroCadena(string nombre, string valor)
        {
            AsignarParametro(nombre, "'", valor);
        }

        /// <summary>
        /// Asigna un parámetro de tipo entero al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroEntero(string nombre, int valor)
        {
            AsignarParametro(nombre, "", valor.ToString());
        }

        /// <summary>
        /// Asigna un parámetro de tipo numeric al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroNumeric(string nombre, decimal valor)
        {
            AsignarParametro(nombre, "", valor.ToString());
        }

        /// <summary>
        /// Asigna un parámetro de tipo fecha al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroFecha(string nombre, DateTime valor)
        {
            AsignarParametro(nombre, "'", valor.ToString("yyyy-MM-dd HH:mm:ss"));
        }

        /// <summary>
        /// Setea un parámetro como nulo del comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro cuyo valor será nulo.</param>
        public void AsignarParametroNulo(string nombre)
        {
            AsignarParametro(nombre, "", "NULL");
        }

        #endregion

        #region Ejecutar
        /// <summary>
        /// Ejecuta el comando creado y retorna el resultado de la consulta.
        /// </summary>
        /// <returns>El resultado de la consulta.</returns>
        /// <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
        public DbDataReader EjecutarConsulta()
        {
            return comando.ExecuteReader();
        }

        /// <summary>
        /// Ejecuta el comando creado y retorna un escalar.
        /// </summary>
        /// <returns>El escalar que es el resultado del comando.</returns>
        /// <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
        public int EjecutarEscalar()
        {
            int escalar = 0;
            try
            {
                escalar = int.Parse(comando.ExecuteScalar().ToString());
            }
            catch (InvalidCastException ex)
            {
                throw new BDException("Error al ejecutar un escalar.", ex);
            }
            return escalar;
        }

        /// <summary>
        /// Ejecuta el comando creado.
        /// </summary>
        public void EjecutarComando()
        {
            comando.ExecuteNonQuery();
        }
        #endregion

        #region Transacciones
        /// <summary>
        /// Comienza una transacción en base a la conexion abierta.
        /// Todo lo que se ejecute luego de esta ionvocación estará 
        /// dentro de una tranasacción.
        /// </summary>
        public void ComenzarTransaccion()
        {
            if (transaccion == null)
            {
                transaccion = conexion.BeginTransaction();
            }
        }

        /// <summary>
        /// Cancela la ejecución de una transacción.
        /// Todo lo ejecutado entre ésta invocación y su 
        /// correspondiente <c>ComenzarTransaccion</c> será perdido.
        /// </summary>
        public void CancelarTransaccion()
        {
            if (transaccion != null)
            {
                transaccion.Rollback();
            }
        }

        /// <summary>
        /// Confirma todo los comandos ejecutados entre el <c>ComanzarTransaccion</c>
        /// y ésta invocación.
        /// </summary>
        public void ConfirmarTransaccion()
        {
            if (transaccion != null)
            {
                transaccion.Commit();
            }
        }
        #endregion

        #region DataReaderToDataSet
        public DataSet DataReader_DataSet()
        {
            DbDataReader reader = EjecutarConsulta();
            DataSet ds = new DataSet();

            try
            {
                do
                {
                    // Create new data table
                    DataTable schemaTable = reader.GetSchemaTable();
                    DataTable dataTable = new DataTable();

                    if (schemaTable != null)
                    {
                        // A query returning records was executed
                        for (int i = 0; i < schemaTable.Rows.Count; i++)
                        {
                            DataRow dataRow = schemaTable.Rows[i];
                            // Create a column name that is unique in the data table
                            string columnName = (string)dataRow["ColumnName"]; //+ "<C" + i + "/>";
                            // Add the column definition to the data table
                            DataColumn column = new DataColumn(columnName, (Type)dataRow["DataType"]);
                            dataTable.Columns.Add(column);
                        }

                        ds.Tables.Add(dataTable);
                        // Fill the data table we just created
                        while (reader.Read())
                        {
                            DataRow dataRow = dataTable.NewRow();
                            for (int i = 0; i < reader.FieldCount; i++)
                                dataRow[i] = reader.GetValue(i);
                            dataTable.Rows.Add(dataRow);
                        }
                    }
                    else
                    {
                        // No records were returned
                        DataColumn column = new DataColumn("RowsAffected");
                        dataTable.Columns.Add(column);
                        ds.Tables.Add(dataTable);
                        DataRow dataRow = dataTable.NewRow();
                        dataRow[0] = reader.RecordsAffected;
                        dataTable.Rows.Add(dataRow);
                    }
                }
                while (reader.NextResult());

            }
            catch { }
            return ds;
        }
        #endregion
    }
}
