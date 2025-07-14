using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Helpers.Files
{
    public class ManagementFilesSD
    {
        private string _path;
        
        public ManagementFilesSD(string path)
        {
            this._path = path;
        }

        public string[] getFiles()
        {
            if (Directory.Exists(this._path))
            {
                string[] files = Directory.GetFiles(this._path);
                return files;
            }
            else
            {
                return null;
            }
        }
    }
}
