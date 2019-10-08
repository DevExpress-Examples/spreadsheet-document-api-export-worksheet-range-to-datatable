Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Export
Imports System.Data
Imports System.Windows.Forms

Namespace ExportToDataTableWorkbookExample
    Partial Public Class Form1
        Inherits DevExpress.XtraBars.Ribbon.RibbonForm

        Private ds As DataTable
        Public Sub New()
            InitializeComponent()

        End Sub

        Private Sub barBtnExport_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barBtnExport.ItemClick
            If ds IsNot Nothing Then
                Return
            End If
'            #Region "#exportdatatable"
            Dim wbook As New Workbook()
            wbook.LoadDocument("TopTradingPartners.xlsx")
            Dim worksheet As Worksheet = wbook.Worksheets(0)
            Dim range As CellRange = worksheet.Tables(0).Range

            Dim dataTable As DataTable = worksheet.CreateDataTable(range, True)
            ' Change the data type of the "As Of" column to text.
            dataTable.Columns("As Of").DataType = System.Type.GetType("System.String")

            Dim exporter As DataTableExporter = worksheet.CreateDataTableExporter(range, dataTable, True)
            AddHandler exporter.CellValueConversionError, AddressOf exporter_CellValueConversionError
            Dim myconverter As New MyConverter()
            exporter.Options.CustomConverters.Add("As Of", myconverter)
            ' Set the export value for empty cell.
            myconverter.EmptyCellValue = "N/A"
            exporter.Options.ConvertEmptyCells = True

            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = False

            exporter.Export()
'            #End Region ' #exportdatatable
            ds = dataTable
            gridControl1.DataSource = ds
        End Sub
        #Region "#CellValueConversionError"
        Private Sub exporter_CellValueConversionError(ByVal sender As Object, ByVal e As CellValueConversionErrorEventArgs)
            MessageBox.Show("Error in cell " & e.Cell.GetReferenceA1())
            e.DataTableValue = Nothing
            e.Action = DataTableExporterAction.Continue
        End Sub
        #End Region ' #CellValueConversionError
    End Class
End Namespace
