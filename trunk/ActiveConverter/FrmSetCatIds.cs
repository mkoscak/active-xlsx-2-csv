using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ActiveConverter
{
    public partial class FrmSetCatIds : Form
    {
        public FrmSetCatIds()
        {
            InitializeComponent();
        }

        public void SetListData(List<DataItem> data)
        {
            gridData.DataSource = data.Select(di => new CatIdDataItem(di)).ToList();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!chkUseDefault.Checked)
                DialogResult = DialogResult.OK;

            Close();
        }
    }

    public class DoubleBufferedGrid : DataGridView
    {
        public DoubleBufferedGrid()
        {
            this.DoubleBuffered = true;
        }
    }
}
