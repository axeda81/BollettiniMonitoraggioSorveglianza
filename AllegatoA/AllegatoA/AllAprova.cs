using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//using PdfSharp.Pdf;
//using PdfSharp.Pdf.AcroForms;
//using PdfSharp.Pdf.Advanced;
//using PdfSharp.Pdf.IO;

//using Gnostice.PDFOne;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace BollettiniMonitoraggio
{
    class AllAprova
    {

        public AllAprova()
        {

            //variables
            String pathin = "AllegatoA.pdf";
            String pathout = "prova.pdf";
            //create a document object
            //var doc = new Document(PageSize.A4);
            //create PdfReader object to read from the existing document
            PdfReader reader = new PdfReader(pathin);
            //select three pages from the original document
            reader.SelectPages("1-2");
            //create PdfStamper object to write to get the pages from reader 
            PdfStamper stamper = new PdfStamper(reader, new FileStream(pathout, FileMode.Create));
            // PdfContentByte from stamper to add content to the pages over the original content
            
            PdfContentByte pbover = stamper.GetOverContent(1);
            PdfWriter writer = stamper.Writer;
            //PdfWriter writer=PdfWriter.GetInstance(doc, new FileStream(path + "/pdfdoc.pdf", FileMode.Create));
            TextField textname = new TextField(writer, new Rectangle(36, 800, 136, 780), "txtname");
            textname.Text = "Enter your name...";
            textname.TextColor = new BaseColor(255, 0, 0);
            textname.BackgroundColor = BaseColor.LIGHT_GRAY;
            //writer.AddAnnotation(textname.GetTextField());
            stamper.AddAnnotation(textname.GetTextField(),2);



            //add content to the page using ColumnText
            //ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase("Hello World"), 100, 400, 0);
            // PdfContentByte from stamper to add content to the pages under the original content
            //PdfContentByte pbunder = stamper.GetUnderContent(1);


            //close the stamper
            stamper.Close();




        }

    }
}
