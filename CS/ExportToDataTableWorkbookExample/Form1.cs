using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using System.Data;
using System.Windows.Forms;

namespace ExportToDataTableWorkbookExample {
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm {
        DataTable dataSource;
        public Form1() {
            InitializeComponent();           
        }

        private void barBtnExport_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            if (dataSource != null) return;
            #region #exportdatatable
            Workbook workbook = new Workbook();
            workbook.LoadDocument("TopTradingPartners.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            CellRange range = worksheet.Tables[0].Range;

            DataTable dataTable = worksheet.CreateDataTable(range, true);
            // Change the data type of the "As Of" column to text.
            dataTable.Columns["As Of"].DataType = System.Type.GetType("System.String");
           
            //Create a DataTable exporter
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
            exporter.CellValueConversionError += exporter_CellValueConversionError;
            MyConverter myconverter = new MyConverter();
            exporter.Options.CustomConverters.Add("As Of", myconverter);

            // Set the export value for empty cell.
            myconverter.EmptyCellValue = "N/A";
            exporter.Options.ConvertEmptyCells = true;
            
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = false;
            //Export data
            exporter.Export();
            #endregion #exportdatatable

            dataSource = dataTable;
            gridControl1.DataSource = dataSource;
        }
        #region #CellValueConversionError
        void exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e) {
            MessageBox.Show("Error in cell " + e.Cell.GetReferenceA1());
            e.DataTableValue = null;
            e.Action = DataTableExporterAction.Continue;
        }
        #endregion #CellValueConversionError
    }
}
