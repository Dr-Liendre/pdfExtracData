using AtexisTool.pdfExtract.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AtexisTool.pdfExtract.ViewModels
{

    public class FilesViewModel:FileModel
    {
        private List<FileModel> _fileListModel;

        public List<FileModel> FileListModel
        {
            get { return _fileListModel; }
            set { _fileListModel = value; }
        }
        public FilesViewModel()
        {
            _fileListModel = new List<FileModel>();
        }               
    }
}
