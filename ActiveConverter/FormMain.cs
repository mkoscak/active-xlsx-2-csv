using System;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace ActiveConverter
{
    public partial class frmMain : Form
    {
        string[] Sources = { "Adidas SLVR", 
                               "Room31 WholeSale" };

        public frmMain()
        {
            InitializeComponent();

            cbExcelType.DataSource = Sources;
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel 2007 files|*.xlsx";
            ofd.CheckFileExists = true;
            ofd.InitialDirectory = Application.StartupPath;
            ofd.Multiselect = false;
            var res = ofd.ShowDialog();
            if (res != DialogResult.OK)
                return;

            this.Cursor = Cursors.WaitCursor;
            Process(ofd.FileName);
            this.Cursor = Cursors.Default;

            MessageBox.Show(this, "Spracovane!", "Ok", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void Process(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            ExcelPackage ep = new ExcelPackage(fi);

            Model myModel = new Model(ep.Workbook.Worksheets[1], txtCatIds.Text ?? string.Empty);

            if (myModel.FindHeaderIndex())
            {
                if (myModel.FindColumnIndices())
                {
                    while (myModel.FindFirstCode())
                    {
                        // ADIDAS SLVR
                        if (cbExcelType.SelectedIndex == 0)
                        {
                            // Create data item and add it to list of outputs
                            if (!myModel.ProcessNextDataItemAdidasSLVR())
                            {
                                MessageBox.Show(this, "Nepodarilo sa nacitat polozku (riadok " + myModel.actualRowIndex + ")!", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        // ROOM31 WHOLESALE
                        else if (cbExcelType.SelectedIndex == 1)
                        {
                            // Create data item and add it to list of outputs
                            if (!myModel.ProcessNextDataItemRoom31())
                            {
                                MessageBox.Show(this, "Nepodarilo sa nacitat polozku (riadok " + myModel.actualRowIndex + ")!", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show(this, "Nepodarilo sa najst nazvy hlaviciek (pictures, code, description, rrp a supp)!", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!myModel.StoreData(Application.StartupPath + @"\exported"))
                {
                    MessageBox.Show(this, "Nepodarilo sa ulozit vystupne data!", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show(this, "Nepodarilo sa najst hlavickovy riadok (aspon 5 zadanych hodnot v stlpcoch za sebou)!", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}
