using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AtexisTool.pdfExtract.Models
{
    public class FileModel
    {
        public int id { get; set; }        
        public string filename { get; set; }
        public string status { get; set; }
        public string remarks { get; set; }        
    }
}
