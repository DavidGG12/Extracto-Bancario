using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.DataAcces
{
    public class BDException : ApplicationException
    {
        /// <summary>
        /// Construye una instancia en base a un mensaje de error y la una excepción original.
        /// </summary>
        /// <param name="mensaje">El mensaje de error.</param>
        /// <param name="original">La excepción original.</param>
        public BDException(string mensaje, Exception original) : base(mensaje, original)
        { }

        /// <summary>
        /// Construye una instancia en base a un mensaje de error.
        /// </summary>
        /// <param name="mensaje">El mensaje de error.</param>
        public BDException(string mensaje) : base(mensaje)
        { }
    }
}
