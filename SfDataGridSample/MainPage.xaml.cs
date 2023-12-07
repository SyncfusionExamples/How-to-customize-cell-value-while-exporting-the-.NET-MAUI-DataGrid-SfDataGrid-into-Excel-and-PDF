using SfDataGridSample.Services;
using Syncfusion.Maui.DataGrid;
using Syncfusion.Maui.DataGrid.Exporting;
using Syncfusion.Maui.DataGrid.Helper;

namespace SfDataGridSample
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void Button_Clicked(object sender, EventArgs e)
        {
            DataGridExcelExportingController excelExport = new DataGridExcelExportingController();
            DataGridExcelExportingOption exportOption = new DataGridExcelExportingOption();
            excelExport.CellExporting += ExcelExport_CellExporting;
            var excelEngine = excelExport.ExportToExcel(this.dataGrid, exportOption);
            var workbook = excelEngine.Excel.Workbooks[0];
            MemoryStream stream = new MemoryStream();
            workbook.SaveAs(stream);
            workbook.Close();
            excelEngine.Dispose();

            string OutputFilename = "DefaultDataGrid.xlsx";
            SaveService saveService = new();
            saveService.SaveAndView(OutputFilename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", stream);
        }

        private void ExcelExport_CellExporting(object? sender, DataGridCellExcelExportingEventArgs e)
        {
            if (e.CellValue is string && e.CellValue == "Phyllis Allen")
            {
                e.CellValue = "Guptill Martin";
            }
        }

        private void Button_Clicked_1(object sender, EventArgs e)
        {
            MemoryStream stream = new MemoryStream();
            DataGridPdfExportingController pdfExport = new DataGridPdfExportingController();
            DataGridPdfExportingOption option = new DataGridPdfExportingOption();
            pdfExport.CellExporting += PdfExport_CellExporting;
            var pdfDoc = pdfExport.ExportToPdf(this.dataGrid, option);
            pdfDoc.Save(stream);
            pdfDoc.Close(true);
            SaveService saveService = new();
            saveService.SaveAndView("ExportFeature.pdf", "application/pdf", stream);
        }

        private void PdfExport_CellExporting(object? sender, DataGridCellPdfExportingEventArgs e)
        {
            if (e.CellValue is string && e.CellValue == "Hannah Arakawa")
            {
                e.CellValue = "Guptill Martin";
            }
        }
    }
}
