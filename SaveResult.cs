using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using System.Windows.Documents;
using iTextSharp.text;
using System.IO;
using System.Threading;
using System.Drawing;



namespace Lyo3
{
    class SaveResult
    {
        DataGridView _dg;
        private List<string> _headers = new List<string>();
        private List<DataGridViewRow> _data;
        int _columnCount;
        string _fileLocation;
        string _fileType;
        string _fileName;

        public String FilterType { get; set; }
        public SaveResult(DataGridView dgObject)
        {
            _dg = dgObject;
            chooseFileLocation();
        }
        private void chooseFileLocation()
        {
            SaveFileDialog sfd = new SaveFileDialog();

            switch (_fileType)
            {
                case "pdf":
                    sfd.DefaultExt = "pdf";
                    break;
                case "xls":
                    sfd.DefaultExt = "xls";
                    break;
            }
            
            DialogResult result = sfd.ShowDialog();

            if(result == DialogResult.OK)
            {
                _fileLocation = sfd.FileName;
            }
        }
        public void StartNewThread(string fileType)
        {
            ThreadStart saveThreadStartPdf = new ThreadStart(SaveToPdf);
            ThreadStart saveThreadStartXls = new ThreadStart(SaveToCSV);
            Thread saveThreadPdf = new Thread(saveThreadStartPdf);
            Thread saveThreadXls = new Thread(saveThreadStartXls);
            switch (fileType)
            {
                case "pdf":
                    saveThreadPdf.Start();
                    break;

                case "xls":
                    saveThreadXls.Start();
                    break;
            }
           
        }
        private void GetDataRows()
        {
            foreach (DataGridViewRow r in _dg.Rows)
            {
                _data.Add(r);
            }
        }
        private void GetDataHeaders()
        {
            foreach (DataGridViewColumn c in _dg.Columns)
            {
                _headers.Add(c.Name);
            }
        }

        public void SaveToCSV()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelWrkBk = excelApp.Workbooks.Add() as Workbook;
            Worksheet excelWrkSht = excelWrkBk.Sheets[1] as Worksheet;

            excelWrkSht.Cells[1, 1].Activate();
            GetDataHeaders();

            for(int i = 1; i <= _headers.Count;i++)
            {
                excelWrkSht.Cells[1,i] = _headers[i - 1];
                excelWrkSht.Cells[1,i].Borders.Weight = XlBorderWeight.xlThick;
            }

            for (int x = 0; x <= _dg.Rows.Count-1; x++)
            {
                for (int y = 0; y <= _dg.Columns.Count-1; y++)
                {
                    excelWrkSht.Cells[x+2, y+1] = _dg[y,x].Value.ToString();
                }
            }
             

            excelApp.Columns.AutoFit();
            excelWrkBk.SaveAs(_fileLocation);
            excelApp.Visible = true;
        }

        public void SaveToPdf()
        {
            

            using (FileStream stream = new FileStream(_fileLocation, FileMode.Create))
            {
                PdfPTable pdfT = new PdfPTable(_dg.Columns.Count);
                Document doc = new Document();
                PdfWriter.GetInstance(doc, stream);
                int count = 2;

                doc.SetPageSize(PageSize.A4.Rotate());
                
                pdfT.WidthPercentage = 100;
                pdfT.HorizontalAlignment = Element.ALIGN_LEFT;
                pdfT.DefaultCell.BorderWidth = 1;
                GetDataHeaders();

                foreach (string r in _headers)
                {
                    PdfPCell cell = new PdfPCell(new iTextSharp.text.Phrase(r));
                    cell.Colspan = 3;
                    cell.Rowspan = 1;
                    cell.BackgroundColor = BaseColor.GRAY;
                    pdfT.AddCell(cell);
                }
                foreach (DataGridViewRow row in _dg.Rows)
                {
                    foreach (DataGridViewCell c in row.Cells)
                    {

                        PdfPCell cell = new PdfPCell(new iTextSharp.text.Phrase(c.Value.ToString()));
                        cell.Rowspan = 1;
                        cell.Colspan = 3;

                        if (count%2 != 0)
                        {
                            cell.BackgroundColor = BaseColor.GRAY;
                        }
                        pdfT.AddCell(cell);
                    }
                    count += 1;

                }
                doc.Open();
                doc.Add(pdfT);
                doc.Close();
                stream.Close();
                
            }
                       
        }
        public System.Drawing.Printing.PrintPageEventArgs Print(System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap _bm = new Bitmap(_dg.Width, _dg.Height);
            _dg.DrawToBitmap(_bm, new System.Drawing.Rectangle(0, 0, _dg.Width, _dg.Height));
            e.PageSettings.Landscape = true;
            e.PageSettings.Color = false;
            e.Graphics.DrawImage(_bm, 0, 0);
            return e;
        }
    }
}
