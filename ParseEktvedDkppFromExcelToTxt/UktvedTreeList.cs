using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace ParserExcelUKTVEDtoBunariList
{
    public partial class UktvedTreeList : DevExpress.XtraEditors.XtraForm
    {
        private BindingSource treeBS = new BindingSource();

        public UktvedTreeList(List<ExcelUKTVEDDTO> listOfCode)
        {
            InitializeComponent();



            //List<ExcelDTO> listAll = listOfCode.GetRange(0, 6000);
            //treeBS.DataSource = listAll;

            treeBS.DataSource = listOfCode;
            treeList.DataSource = treeBS;
            treeList.KeyFieldName = "Id";
            treeList.ParentFieldName = "ParentId";
            //treeList.ExpandAll();
        }
    }
}