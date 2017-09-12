using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.Advanced;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace BollettiniMonitoraggio
{
    class AllA
    {
        readonly static Color azzurro1 = new Color(171, 205, 239);
        readonly static Color azzurro2 = new Color(0, 127, 255);
        readonly static Color azzurro3 = new Color(0, 0, 128);
        readonly static Color grigio = Colors.Gray;
        readonly static Color grigioChiaro = Colors.LightGray;
        readonly static Color verde = Colors.LawnGreen;
        readonly static Color giallo = Colors.LightYellow;

        protected MigraDoc.DocumentObjectModel.Document document;
        protected MigraDoc.DocumentObjectModel.Section section;
        protected PdfDocumentRenderer pdfRenderer;
        protected StreamReader fileCSV1;
        protected StreamReader fileCSV3;
        protected string nomeFileCSV1;
        protected string nomeFileCSV3;
        protected string indiceBollettino = "";
        protected string numAvviso = "";
        protected string dataAvviso = "";
        protected string dataInizio = "";
        protected string oraInizio = "";
        protected string dataFine = "";
        protected string oraFine = "";
        protected Table tableAll1;
        protected Table tableAll3;
        private LogWriter log; // collegamento al file di log dove scriverò informazioni sull'esecuzione
        protected string filename; // nome file pdf
        
        protected int altezzaElementi;
        protected int numPagina;

        const int numColonneAll1 = 13;

        /* Costruttore AllA
         * all1: stringa che contiene il nome del file CSV con i dati dell'all. 1
         * all3: stringa che contiene il nome del file CSV con i dati dell'all. 3
         * i: stringa che contiene il numero progressivo del bollettino nel formato xxx/aaaa
         * avviso: stringa che contiene numero e data dell'avviso di criticità
         * inizio: stringa con ora e data dell'inizio della validità dell'avviso
         * fine: stringa con ora e data della fine della validità dell'avviso
         */
        public AllA(string all1, string all3, string i, string n_avv, string data_avv, string data_i, string ora_i, string data_f, string ora_f)
        {
            log = LogWriter.Instance; // File di log in cui scrivo informazioni su esecuzione ed errori
            string consultazione = "Consultazione alle ore " + DateTime.Now.ToString("HH:mm") + " del " + DateTime.Now.ToString("dd/MM/yyyy");
            log.WriteToLog("ALL. A - Inizio scrittura bollettino, " + consultazione + ".", level.Info);
            //filename = DateTime.Now.ToString("ALL. A - ddMMyyyy-HHmm_") + indiceBollettino + ".pdf";
            filename = "AllegatoA.pdf";
            altezzaElementi = 0;
            numPagina = 0;
            nomeFileCSV1 = all1;
            nomeFileCSV3 = all3;
            indiceBollettino = i;
            numAvviso = n_avv;
            dataAvviso = data_avv;
            dataInizio = data_i;
            oraInizio = ora_i;
            dataFine = data_f;
            oraFine = ora_f;
            CreaPDF();
            log.WriteToLog("ALL. A - Fine scrittura bollettino, " + consultazione + ".", level.Info);
        }

        private void CreaPDF()
        {
            try
            {
                // Apro i due CSV che mi servono per riempire l'All.A
                using (fileCSV1 = new StreamReader(nomeFileCSV1))
                using(fileCSV3 = new StreamReader(nomeFileCSV3)) 
                {
                    // Creazione del PDF
                    document = new MigraDoc.DocumentObjectModel.Document();
                    document.Info.Title = "Bollettino di monitoraggio e sorveglianza";
                    document.Info.Author = "R.A.S. Protezione Civile";
                    document.UseCmykColor = true;
                    const bool unicode = true;
                    const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

                    // Creo o sovrascrive il file pdf
                    File.Create(filename);

                    DefineStyles();
                    // Impaginazione delle informazioni nel pdf
                    CreatePage();

                    pdfRenderer = new PdfDocumentRenderer(unicode, embedding); // opzioni di visualizzazione
                    pdfRenderer.Document = document;
                    pdfRenderer.RenderDocument();
                 
                    // Alla fine salva il pdf appena creato: se esiste già con lo stesso nome lo elimina e poi salva quello nuovo
                    if (File.Exists(filename))
                    {
                        FileInfo bInfoOld = new FileInfo(filename);
                        bInfoOld.IsReadOnly = false;
                        File.Delete(filename);
                    }

                    File.Create(filename);
                    pdfRenderer.PdfDocument.Save(filename);
                 
                    // Rendo il pdf finale modificabile
                    if (File.Exists(filename))
                    {
                        FileInfo bInfoOld = new FileInfo(filename);
                        bInfoOld.IsReadOnly = false;
                    }

                    //System.Diagnostics.Process.Start(filename);

                } // using (fileCSV1, fileCSV3)
            }//try

            catch (System.Exception err)
            {
                // Intercetta tutti i tipi di eccezione ma quelle che si dovrebbero verificare più di frequente sono:
                // System.IO.FileNotFoundException, System.IO.DirectoryNotFoundException, System.IO.IOException...        
                err.Source = "CreaPDF()";
                // Salvo nel log il messaggio di errore con un pò di informazioni sulla funzione che ha lanciato l'eccezione e sul tipo di eccezione
                log.WriteToLog("ALL. A: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception);
            }
        }

        void CreatePage()
        {
            // Each MigraDoc document needs at least one section.
            section = document.AddSection();
            // Il pdf verrà stampato in orizzontale
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Portrait;
            section.PageSetup.TopMargin = "05mm";
            section.PageSetup.LeftMargin = "05mm";
            section.PageSetup.RightMargin = "05mm";
            section.PageSetup.BottomMargin = "15mm";
            section.PageSetup.DifferentFirstPageHeaderFooter = false;
            section.PageSetup.FooterDistance = Unit.FromCentimeter(0.2);

            // ******************* Da qui in poi - creazione del contenuto del PDF *******************

            CreatePageHeader(); // Header 

            CreateContent(); // Dati bollettini

            CreateFooter(); // Footer 

        } //CreatePage()   

        void CreatePageHeader()
        {
            // Creo l'header del documento 
            Table tableHeader = section.AddTable(); 
            tableHeader.Borders.Width = 0.2;
            tableHeader.Rows.LeftIndent = 0;
            tableHeader.Format.Alignment = ParagraphAlignment.Center;

            // Definisco le colonne
            Column column = tableHeader.AddColumn("1cm");
            column = tableHeader.AddColumn("3cm");
            column = tableHeader.AddColumn("3cm");
            column = tableHeader.AddColumn("3cm");
            column = tableHeader.AddColumn("3cm");
            column = tableHeader.AddColumn("3cm");
            column = tableHeader.AddColumn("1cm");
            column = tableHeader.AddColumn("1cm");
            column = tableHeader.AddColumn("2cm");

            double dimImg = 2;
            Row row = tableHeader.AddRow();
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].Borders.Right.Visible = false;
            
            MigraDoc.DocumentObjectModel.Shapes.Image image = new MigraDoc.DocumentObjectModel.Shapes.Image("./LOGO RAS.tif");
            image.LockAspectRatio = true;
            image.Height = Unit.FromCentimeter(dimImg);
            image.Width = Unit.FromCentimeter(dimImg);
            row.Cells[1].Borders.Left.Visible = false;
            row.Cells[1].Add(image);

            row.Cells[2].MergeRight = 3;
            row.Cells[7].MergeRight = 1;
            row.Cells[6].Borders.Right.Visible = false;
            row.Cells[7].Borders.Left.Visible = false;

            row.Cells[2].Shading.Color = Colors.LightGray;
            row.Cells[2].AddParagraph("Centro Funzionale Decentrato Regione Sardegna").Format.Font.Size = 10;
            row.Cells[2].AddParagraph("BOLLETTINO DI MONITORAGGIO").Format.Font.Size = 16;
            row.Cells[2].Format.Font.Bold = true;

            MigraDoc.DocumentObjectModel.Shapes.Image imagePC = new MigraDoc.DocumentObjectModel.Shapes.Image("./LOGO CIRCOLARE.gif");
            imagePC.LockAspectRatio = true;
            imagePC.Height = Unit.FromCentimeter(dimImg);
            imagePC.Width = Unit.FromCentimeter(dimImg);
            row.Cells[7].Add(imagePC);

            // In questa riga devono essere inseriti i dati relativi all'avviso di criticità in corso
            row = tableHeader.AddRow();
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Height = 18;
            row.Cells[0].MergeRight = 2;
            row.Cells[0].AddParagraph("Avviso di criticità n. " + numAvviso + " del " + dataAvviso);
            row.Cells[3].AddParagraph("Inizio validità");
            row.Cells[4].AddParagraph(oraInizio + " del " + dataInizio);
            row.Cells[5].AddParagraph("Fine validità");
            row.Cells[6].MergeRight = 2;
            row.Cells[6].AddParagraph(oraFine + " del " + dataFine);

            // Riga di spazio
            row = tableHeader.AddRow();
            row.Height = 3;
            row.Borders.Visible = false;

            // Informazioni su data e ora di emissione del bollettino
            row = tableHeader.AddRow();
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Height = 18;
            row.Cells[0].MergeRight = 1;
            row.Cells[0].AddParagraph("Numero progressivo");
            row.Cells[2].AddParagraph(indiceBollettino);
            row.Cells[3].AddParagraph("Data di emissione");
            row.Cells[4].AddParagraph(DateTime.Now.ToString("dd/MM/yyyy"));
            row.Cells[5].AddParagraph("Ora locale");
            row.Cells[6].MergeRight = 2;
            row.Cells[6].AddParagraph(DateTime.Now.ToString("HH:mm"));

            
        }

        private void CreateContent()
        {
            /* Contenuto del documento: 
            * - tabella pluviometri + legenda
            * - tabella idrometri + legenda
            * - commento settore idro
            * - valutazione meteorologica settore meteo
            * - valutazioni idrauliche settore idro
            */
            MigraDoc.DocumentObjectModel.Paragraph paragraph = section.AddParagraph("\n\nAnalisi dei dati pluviometrici e idrometrici della rete fiduciaria di protezione civile\n\n");
            paragraph.Format.Font.Bold = true;
            paragraph.Format.Font.Size = 10;
            paragraph.Format.Alignment = ParagraphAlignment.Left;

            MigraDoc.DocumentObjectModel.Paragraph nota = section.AddParagraph("\"Composizione e rappresentazione dei dati eseguita con modalità automatiche su dati della rete di stazioni meteorologiche fiduciarie della Regione Sardegna gestita dall\'Agenzia per la Protezione dell'Ambiente della Sardegna, ARPAS, acquisiti in tempo reale e sottoposti ad un processo automatico di validazione di primo livello\"");
            nota.Format.Font.Bold = false;
            nota.Format.Font.Size = 6;
            nota.Format.Alignment = ParagraphAlignment.Center;

            // Riempimento tabella con dati dei pluviometri che hanno superato soglie e relativa legenda
            HeaderTableAll1();
            ContentTableAll1();
            LegendaAll1();

            section.AddParagraph().AddLineBreak();
            section.AddParagraph().AddLineBreak();

            // Riempimento tabella con dati degli idrometri che hanno superato soglie e relativa legenda
            HeaderTableAll3();
            ContentTableAll3();
            LegendaAll3();

            // Creazione dei contenitori dei campi compilabili
            CampiCompilabili();

            // Inserimento dell'area per il commento del CFD - settore idro
            //Row row = tableAll3.AddRow();
            //row.Height = 4;
            //row.Cells[0].MergeRight = 10;
            //row.Cells[0].Borders.Visible = false;
            //row.Shading.Color = Colors.White;
       
            //Row row = tableAll3.AddRow();
            //row.Cells[0].MergeRight = 10;
            //row.Cells[0].AddParagraph("Commento (a cura del CFD settore Idro)");
            //row.Cells[0].Format.Font.Bold = true;
            //row.Cells[0].Format.Font.Size = 8;
            //row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            //row.Cells[0].AddParagraph("\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
            //row.Format.Alignment = ParagraphAlignment.Left;
            //row.Cells[0].Format.Font.Size = 7;

            //section.AddParagraph().AddLineBreak();

            /*
            Table commento = section.AddTable();
            commento.Borders.Visible = true;
            commento.Borders.Width = 0.2;
            commento.Rows.LeftIndent = 0;
            commento.AddColumn("20cm");

            Row row = commento.AddRow();
            Paragraph p = row.Cells[0].AddParagraph("Commento\n\n\n\n");
            p.Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Font.Size = 8;
            FormattedText ft = p.AddFormattedText("\t\t\t\t\tCommento testuale\n\n");
            ft.Style = "Comment"; 

            section.AddParagraph().AddLineBreak();

            // Inserimento dell'area per il commento del CFD - settore meteo
            Table meteo = section.AddTable();
            meteo.Borders.Visible = true;
            meteo.Borders.Width = 0.2;
            meteo.Rows.LeftIndent = 0;
            meteo.AddColumn("20cm");
            Row firstrow = meteo.AddRow();
            firstrow.Format.Alignment = ParagraphAlignment.Left;
            firstrow.Cells[0].AddParagraph("Valutazione meteorologica");
            firstrow.Shading.Color = grigio;
            firstrow.Format.Font.Color = Colors.White;
            row = meteo.AddRow();
            row.KeepWith = firstrow.Index;
            p = row.Cells[0].AddParagraph();
            ft = p.AddFormattedText("\n Valutazione testuale\n\n");
            ft.Style = "Comment";
            row = meteo.AddRow();
            row.KeepWith = firstrow.Index;
            row.Shading.Color = grigio;
            row.Cells[0].AddParagraph();
            firstrow.KeepWith = meteo.Rows.Count - 1;

            section.AddParagraph().AddLineBreak();

            // Inserimento dell'area per le valutazioni idrauliche del CFD - settore idro
            Table idro = section.AddTable();
            idro.Borders.Visible = true;
            idro.Borders.Width = 0.2;
            idro.Rows.LeftIndent = 0;
            idro.AddColumn("20cm");

            firstrow = idro.AddRow();
            firstrow.Format.Alignment = ParagraphAlignment.Left;
            firstrow.Cells[0].AddParagraph("Valutazioni idrauliche");
            firstrow.Shading.Color = grigio;
            firstrow.Format.Font.Color = Colors.White;
            row = idro.AddRow();
            row.KeepWith = firstrow.Index;
            p = row.Cells[0].AddParagraph();
            ft = p.AddFormattedText("\n Valutazione testuale\n\n");
            ft.Style = "Comment";
            row = idro.AddRow();
            row.KeepWith = firstrow.Index;
            row.Shading.Color = grigio;
            row.Cells[0].AddParagraph();
            firstrow.KeepWith = idro.Rows.Count - 1;

            section.AddParagraph().AddLineBreak();
            section.AddParagraph("Il direttore del Servizio previsione rischi e dei sistemi informativi, infrastrutture e reti\n\n");
            */
        }

        private void CreateFooter()
        {
            // Aggiungo il footer
            MigraDoc.DocumentObjectModel.Paragraph footer = new MigraDoc.DocumentObjectModel.Paragraph();
            footer.AddLineBreak(); 
            FormattedText f1 = new FormattedText();
            f1.AddFormattedText("Centro funzionale decentrato della Regione Sardegna: via Vittorio Veneto 28, 09128 Cagliari\n");
            f1.Font.Size = 5;
            footer.Add(f1);
            FormattedText f2 = new FormattedText();
            f2.AddFormattedText("cfd.protezionecivile@pec.regione.sardegna.it - protciv.previsioneprevenzionerischi@regione.sardegna.it");
            f2.Font.Size = 4;
            f2.Font.Underline = Underline.Single;
            f2.Font.Color = Colors.Blue;
            footer.Add(f2);

            section.Footers.Primary.Add(footer);
            section.Footers.EvenPage.Add(footer.Clone());

            // Aggiungo i numeri di pagina
            section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            MigraDoc.DocumentObjectModel.Paragraph pageNum = new MigraDoc.DocumentObjectModel.Paragraph();
            pageNum.AddLineBreak();
            pageNum.AddText("Pagina ");
            pageNum.AddPageField();
            pageNum.AddText(" di ");
            pageNum.AddNumPagesField();
            pageNum.Format.Alignment = ParagraphAlignment.Right;
            pageNum.Format.Font.Size = 6;
            section.Footers.Primary.Add(pageNum);
            section.Footers.EvenPage.Add(pageNum.Clone());
        }
        

        private void HeaderTableAll1()
        {
            // TABELLA PLUVIOMETRI
            tableAll1 = section.AddTable();
            tableAll1.Rows.HeightRule = RowHeightRule.Auto;
            tableAll1.Borders.Width = 0.2;
            tableAll1.Rows.LeftIndent = 0;
            tableAll1.Format.Alignment = ParagraphAlignment.Center;
            tableAll1.Rows.VerticalAlignment = VerticalAlignment.Center;
            tableAll1.Rows.Height = 12;
            tableAll1.Format.Font.Size = 6;
            tableAll1.TopPadding = 1;
            tableAll1.BottomPadding = 1;         

            // Definisco le colonne
            Column column = tableAll1.AddColumn("1cm");
            column = tableAll1.AddColumn("1.5cm");
            column = tableAll1.AddColumn("2cm");
            column = tableAll1.AddColumn("2.5cm");
            column = tableAll1.AddColumn("1cm");
            column = tableAll1.AddColumn("1cm");
            column.Shading.Color = giallo;
            column = tableAll1.AddColumn("1cm");
            column.Shading.Color = giallo;
            column = tableAll1.AddColumn("1.5cm");
            column.Shading.Color = giallo;
            column = tableAll1.AddColumn("2.5cm");
            column.Shading.Color = giallo;
            column = tableAll1.AddColumn("2cm");
            column = tableAll1.AddColumn("1cm");
            column = tableAll1.AddColumn("1.5cm");
            column = tableAll1.AddColumn("1.5cm");

            Row row = tableAll1.AddRow();
            row.Cells[0].MergeRight = numColonneAll1 - 1;
            row.Shading.Color = grigio;
            row.Cells[0].AddParagraph("PLUVIOMETRI");
            row.Cells[0].Format.Font.Bold = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].Format.Font.Color = Colors.White;
            row.Format.Font.Size = 7;

            // Intestazione tabella
            row = tableAll1.AddRow();
            row.HeadingFormat = true;
            row.Format.Font.Bold = true;

            row.Cells[0].AddParagraph("N.");
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph("Stazione");
            row.Cells[1].MergeDown = 1;
            row.Cells[2].AddParagraph("Comune");
            row.Cells[2].MergeDown = 1;
            row.Cells[3].AddParagraph("Zona di allerta");
            row.Cells[3].MergeDown = 1;
            row.Cells[4].AddParagraph("Quota (m.s.l.m.)");
            row.Cells[4].MergeDown = 1;
            row.Cells[5].AddParagraph("Pioggia critica di\nriferimento (mm)");
            row.Cells[5].MergeRight = 1;
            row.Cells[7].AddParagraph("Finestra di\nosservazione");
            row.Cells[7].MergeRight = 1;

            row.Cells[9].AddParagraph("Durate di\nprecipitazione Δt");
            row.Cells[9].MergeDown = 1;
            row.Cells[10].AddParagraph("h (mm)");
            row.Cells[10].MergeDown = 1;

            MigraDoc.DocumentObjectModel.Paragraph p = new MigraDoc.DocumentObjectModel.Paragraph();
            p.AddFormattedText("h/h");
            FormattedText ft = p.AddFormattedText("Tr20anni");
            ft.Subscript = true;
            row.Cells[11].Add(p);
            row.Cells[11].MergeDown = 1;

            MigraDoc.DocumentObjectModel.Paragraph p1 = new MigraDoc.DocumentObjectModel.Paragraph();
            p1.AddFormattedText("h/h");
            FormattedText ft1 = p1.AddFormattedText("Tr100anni");
            ft1.Subscript = true;
            row.Cells[12].Add(p1);
            row.Cells[12].MergeDown = 1;

            for (int i = 0; i < numColonneAll1; i++)
            {
                row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
            }

            row = tableAll1.AddRow();
            row.Format.Font.Bold = true;
            MigraDoc.DocumentObjectModel.Paragraph p2 = new MigraDoc.DocumentObjectModel.Paragraph();
            p2.AddFormattedText("h");
            FormattedText ft2 = p2.AddFormattedText("Tr20anni");
            ft2.Subscript = true;
            row.Cells[5].Add(p2);
            MigraDoc.DocumentObjectModel.Paragraph p3 = new MigraDoc.DocumentObjectModel.Paragraph();
            p3.AddFormattedText("h");
            FormattedText ft3 = p3.AddFormattedText("Tr100anni");
            ft3.Subscript = true;
            row.Cells[6].Add(p3);

            row.Cells[7].AddParagraph("dalle ore");
            row.Cells[8].AddParagraph("alle ore");
            for (int i = 5; i < 9; i++)
            {
                row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
            }
        }

        private void ContentTableAll1()
        {
            // Fare prove con csv con formato non compatibile per vedere gestione eccezioni (teoricamente tutte gestite dal chiamante)

            string riga = fileCSV1.ReadLine();
            string[] elencoCampi = { "" };
            bool alternaColoreRighe = true; // Mi serve per colorare una stazione di grigio e una lasciarla bianca, per facilitare la lettura
            const int numRighePerStazione = 5; // sappiamo che per ogni stazione ci sono sempre 5 righe
            int numBacini = 0;
            int indice = -1;
            int[] numStazioniPerBacino; // Array che conterrà a ogni posizione il numero di stazioni per quel bacino
            int[] numRighePerBacino; // Array che conterrà a ogni posizione il numero di righe da inserire nel PDF per quel bacino

            // Prima lettura file: vedo quanti sono i bacini
            while ((riga != null) && (riga != ""))
            {
                elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                if ((elencoCampi[0] != "") && (elencoCampi[1] == ""))
                {
                    // La riga appena letta è l'intestazione di un bacino perchè la prima cella non è vuota
                    // (c'è scritto il nome del bacino), ma quella dopo sì (è sufficiente come controllo?)
                    numBacini++;
                }
                riga = fileCSV1.ReadLine();
            }

            if (numBacini == 0)
            {
                // Il file CSV relativo all'allegato1 era vuoto, lo segnalo nel log
                log.WriteToLog("Allegato A - File allegato1.csv vuoto!", level.Warning);
                Row row = tableAll1.AddRow();
                row.Format.Font.Bold = false;
                row.Cells[0].AddParagraph("Dati non presenti");
                row.Cells[0].MergeRight = numColonneAll1 - 1;
                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            }
            else
            {
                // Scrivo nel file PDF solo se c'erano dati nel CSV
                numStazioniPerBacino = new int[numBacini];
                numRighePerBacino = new int[numBacini];
                fileCSV1.BaseStream.Seek(0, SeekOrigin.Begin); // Torno a inizio file csv

                riga = fileCSV1.ReadLine();
                elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                // Seconda lettura file: quante stazioni per ogni bacino
                while ((riga != null) && (riga != ""))
                {
                    elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                    if ((elencoCampi[0] != "") && (elencoCampi[1] == "")) indice++; // La riga è l'intestazione di un bacino
                    else
                    {
                        if ((elencoCampi[0] != "") && (elencoCampi[1] != ""))
                        {
                            // La riga appena letta è l'intestazione di una stazione perchè i campi non sono vuoti
                            numStazioniPerBacino[indice]++;
                        }

                        // In tutti i casi controllo se la riga contiene info su una stazione che supera una soglia
                        if (((elencoCampi[12] != "") && (Convert.ToDouble(elencoCampi[12]) >= 0.5)) ||
                            ((elencoCampi[13] != "") && (Convert.ToDouble(elencoCampi[13]) >= 0.5)))
                            numRighePerBacino[indice]++;
                    }
                    riga = fileCSV1.ReadLine();
                }

                // ****** A questo punto costruzione del contenuto del PDF con le informazioni appena recuperate ******

                fileCSV1.BaseStream.Seek(0, SeekOrigin.Begin); // Torno a inizio file csv

                for (int b = 0; b < numBacini; b++)
                {
                    if (numRighePerBacino[b] > 0)
                    {
                        // Devo scrivere l'intestazione del bacino solo se ci sono righe per quel bacino da scrivere nel PDF
                        // All'indice i dell'array numStazioniPerBacino ci sarà il numero di stazioni per l'i-esimo bacino
                        riga = fileCSV1.ReadLine(); // La prima riga dovrebbe contenere solo il nome del bacino, lo riporto nel PDF

                        // Elimino il primo carattere del CSV, se è il char di controllo della codifica 
                        if ((b == 0) && (riga.ToCharArray()[0] == 65279))
                            riga = riga.Remove(0, 1);

                        elencoCampi = CSVRowToStringArray(riga, ';', '\n');
                        Row row = tableAll1.AddRow();
                        row.Format.Font.Bold = false;
                        row.Cells[0].Shading.Color = verde;
                        row.Cells[0].AddParagraph(elencoCampi[0]);
                        row.Cells[0].MergeRight = numColonneAll1 - 1;
                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;

                        int righeScrittePerStazione = 0;

                        // Scorro le stazioni di quel bacino perchè ce ne sarà almeno una che ha almeno una riga da riportare nel PDF
                        for (int s = 0; s < numStazioniPerBacino[b]; s++)
                        {
                            // Mi devo salvare le informazioni generali della stazione
                            string[] firstRow = new string[5];

                            // Per ogni stazione nel file ci saranno numRighePerStazione righe, per ognuna devo controllare le soglie
                            for (int r = 0; r < numRighePerStazione; r++)
                            {
                                riga = fileCSV1.ReadLine();
                                elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                                if (r == 0)
                                {
                                    // Informazioni che si ripeteranno in ogni riga che riguarda la stessa stazione
                                    firstRow[0] = elencoCampi[0];
                                    firstRow[1] = elencoCampi[2];
                                    firstRow[2] = elencoCampi[3];
                                    firstRow[3] = elencoCampi[4];
                                    firstRow[4] = elencoCampi[5];
                                }

                                if (((elencoCampi[12] != "") && (Convert.ToDouble(elencoCampi[12]) >= 0.5)) ||
                                    ((elencoCampi[13] != "") && (Convert.ToDouble(elencoCampi[13]) >= 0.5)))
                                {
                                    // OK la riga va riportata nel PDF
                                    row = tableAll1.AddRow();
                                   
                                    row.Cells[0].AddParagraph(firstRow[0]);
                                    row.Cells[1].AddParagraph(firstRow[1]);
                                    row.Cells[2].AddParagraph(firstRow[2]);
                                    row.Cells[3].AddParagraph(firstRow[3]);
                                    row.Cells[4].AddParagraph(firstRow[4]);

                                    for (int c = 5; c < numColonneAll1; c++)
                                    {
                                        row.Cells[c].AddParagraph(elencoCampi[c+1]);
                                    }

                                    for (int f = 0; f < row.Cells.Count; f++)
                                    {
                                        row.Cells[f].VerticalAlignment = VerticalAlignment.Center;
                                        if (alternaColoreRighe) row.Cells[f].Shading.Color = grigioChiaro;
                                    }
                                    row.Cells[5].Shading.Color = row.Cells[6].Shading.Color = row.Cells[7].Shading.Color = row.Cells[8].Shading.Color = giallo;

                                    // Colorazione celle a seconda del livello di superamento soglia
                                    // Colonna h/htr20anni
                                    if (elencoCampi[12] != "")
                                    {
                                        if ((Convert.ToDouble(elencoCampi[12]) >= 0.5) && (Convert.ToDouble(elencoCampi[12]) < 0.75))
                                            row.Cells[11].Shading.Color = azzurro1;
                                        else if ((Convert.ToDouble(elencoCampi[12]) >= 0.75) && (Convert.ToDouble(elencoCampi[12]) < 1))
                                        {
                                            row.Cells[11].Shading.Color = azzurro2;
                                            row.Cells[11].Format.Font.Color = Colors.White;
                                        }

                                        else if (Convert.ToDouble(elencoCampi[12]) >= 1)
                                        {
                                            row.Cells[11].Shading.Color = azzurro3;
                                            row.Cells[11].Format.Font.Color = Colors.White;
                                        }
                                        else
                                            row.Cells[11].Shading.Color = Colors.White;
                                    }
                                    // Colonna h/htr100anni
                                    if (elencoCampi[13] != "")
                                    {
                                        if ((Convert.ToDouble(elencoCampi[13]) >= 0.5) && (Convert.ToDouble(elencoCampi[13]) < 0.75))
                                            row.Cells[12].Shading.Color = azzurro1;
                                        else if ((Convert.ToDouble(elencoCampi[13]) >= 0.75) && (Convert.ToDouble(elencoCampi[13]) < 1))
                                        {
                                            row.Cells[12].Shading.Color = azzurro2;
                                            row.Cells[12].Format.Font.Color = Colors.White;
                                        }
                                        else if (Convert.ToDouble(elencoCampi[13]) >= 1)
                                        {
                                            row.Cells[12].Shading.Color = azzurro3;
                                            row.Cells[12].Format.Font.Color = Colors.White;
                                        }
                                        else
                                            row.Cells[12].Shading.Color = Colors.White;

                                    }
                                    righeScrittePerStazione++;
                                }
                            }

                            // Alla fine dell'esame della stazione, devo unire in verticale le prime celle, che contengono
                            //  nome stazione, comune, ...i campi uguali per tutte le righe che riguardano quella stazione
                            if (righeScrittePerStazione > 1)
                            {
                                Row tmpRow = tableAll1.Rows[row.Index - righeScrittePerStazione + 1];
                                tmpRow.Cells[0].MergeDown = tmpRow.Cells[1].MergeDown = tmpRow.Cells[2].MergeDown = tmpRow.Cells[3].MergeDown =
                                    tmpRow.Cells[4].MergeDown = righeScrittePerStazione - 1;
                            }

                            alternaColoreRighe = !alternaColoreRighe;
                            righeScrittePerStazione = 0;
                        }
                    }
                    else
                    {
                        // Se per un certo bacino non ho inserito righe nel PDF, devo solo saltare tutte le righe del CSV relative a quel bacino
                        // num righe da saltare = 1 con il nome del bacino + numRighePerStazione*quante stazioni ci sono nel bacino
                        int righeDaSaltare = numStazioniPerBacino[b] * numRighePerStazione;
                        for (int salta = 0; salta <= righeDaSaltare; salta++) fileCSV1.ReadLine();
                    }
                }

                // Se nessuna stazione superava la soglia, o se il file CSV era vuoto, aggiungo una riga alla tabella in cui lo scrivo
                int sommaRighe = 0;
                for (int b = 0; b < numBacini; b++)
                    sommaRighe = sommaRighe + numRighePerBacino[b];
                if (sommaRighe == 0)
                {
                    Row row = tableAll1.AddRow();
                    row.Height = 18;
                    row.Cells[0].MergeRight = numColonneAll1 - 1;
                    row.Cells[0].Format.Font.Bold = true;
                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    if (numBacini == 0)
                        row.Cells[0].AddParagraph("Dati non presenti.");
                    else row.Cells[0].AddParagraph("Nessuna stazione supera le soglie. ");
                }

            } // else 

        } // ContentTableAll1


        private void LegendaAll1()
        {
            section.AddParagraph().AddLineBreak();

            Table legenda = section.AddTable();
            legenda.Borders.Width = 0.2;
            legenda.Rows.LeftIndent = 0;
            legenda.Format.Alignment = ParagraphAlignment.Center;
            legenda.Rows.VerticalAlignment = VerticalAlignment.Center;
            legenda.Format.Font.Size = 8;
            legenda.Format.Font.Bold = true;

            legenda.AddColumn("4cm");
            legenda.AddColumn("4cm");
            legenda.AddColumn("4cm");
            legenda.AddColumn("4cm");
            legenda.AddColumn("4cm");

            Row row = legenda.AddRow();
            row.Height = 18;
            row.Cells[0].Shading.Color = grigio;
            row.Cells[0].Format.Font.Color = Colors.Black;
            row.Cells[0].AddParagraph("Legenda dei colori");

            row.Cells[1].Shading.Color = Colors.White;
            MigraDoc.DocumentObjectModel.Paragraph p1 = new MigraDoc.DocumentObjectModel.Paragraph();
            FormattedText ft1 = p1.AddFormattedText("h/h");
            ft1 = p1.AddFormattedText("tr");
            ft1.Subscript = true;
            Text t1 = new Text(" < 50%");
            p1.Clone();
            p1.Add(t1);
            row.Cells[1].Add(p1);

            row.Cells[2].Shading.Color = azzurro1;
            MigraDoc.DocumentObjectModel.Paragraph p2 = new MigraDoc.DocumentObjectModel.Paragraph();
            FormattedText ft2 = p2.AddFormattedText("50% ≤ h/h");
            ft2 = p2.AddFormattedText("tr");
            ft2.Subscript = true;
            Text t2 = new Text(" < 75%");
            p2.Clone();
            p2.Add(t2);
            row.Cells[2].Add(p2);

            row.Cells[3].Shading.Color = azzurro2;
            row.Cells[3].Format.Font.Color = Colors.White;
            MigraDoc.DocumentObjectModel.Paragraph p3 = new MigraDoc.DocumentObjectModel.Paragraph();
            FormattedText ft3 = p3.AddFormattedText("75% ≤ h/h");
            ft3 = p3.AddFormattedText("tr");
            ft3.Subscript = true;
            Text t3 = new Text(" < 100%");
            p3.Clone();
            p3.Add(t3);
            row.Cells[3].Add(p3);

            row.Cells[4].Shading.Color = azzurro3;
            row.Cells[4].Format.Font.Color = Colors.White;
            MigraDoc.DocumentObjectModel.Paragraph p4 = new MigraDoc.DocumentObjectModel.Paragraph();
            FormattedText ft4 = p4.AddFormattedText("h/h");
            ft4 = p4.AddFormattedText("tr");
            ft4.Subscript = true;
            Text t4 = new Text(" ≥ 100%");
            p1.Clone();
            p4.Add(t4);
            row.Cells[4].Add(p4);
        }

        private void HeaderTableAll3()
        {
            // TABELLA IDROMETRI
            tableAll3 = section.AddTable();
            tableAll3.Borders.Width = 0.2;
            tableAll3.Rows.LeftIndent = 0;
            tableAll3.Format.Alignment = ParagraphAlignment.Center;
            tableAll3.Rows.VerticalAlignment = VerticalAlignment.Center;
            tableAll3.Rows.Height = 12;
            tableAll3.Format.Font.Size = 6;
            tableAll3.TopPadding = 1;
            tableAll3.BottomPadding = 1;
            tableAll3.KeepTogether = true;
            
            // Definisco le colonne
            Column column = tableAll3.AddColumn("2.5cm");
            column = tableAll3.AddColumn("2cm");
            column = tableAll3.AddColumn("2.5cm");
            column = tableAll3.AddColumn("2.5cm");
            column = tableAll3.AddColumn("2.5cm");
            column = tableAll3.AddColumn("1.3cm");
            column = tableAll3.AddColumn("0.8cm");
            column.Shading.Color = giallo;
            column = tableAll3.AddColumn("0.8cm");
            column.Shading.Color = giallo;
            column = tableAll3.AddColumn("0.8cm");
            column.Shading.Color = giallo;
            column = tableAll3.AddColumn("1.3cm");
            column = tableAll3.AddColumn("3cm");

            Row row = tableAll3.AddRow();
            row.Cells[0].MergeRight = 10;
            row.Shading.Color = grigio;
            row.Cells[0].AddParagraph("IDROMETRI");
            row.Cells[0].Format.Font.Bold = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].Format.Font.Color = Colors.White;
            row.Format.Font.Size = 7;

            row = tableAll3.AddRow();
            // Intestazione tabella
            row.Borders.Visible = true;
            row.HeadingFormat = true;
            row.Format.Font.Bold = true;

            row.Cells[0].AddParagraph("Stazione");
            row.Cells[1].AddParagraph("Comune");
            row.Cells[2].AddParagraph("Zona di allerta");
            row.Cells[3].AddParagraph("Bacino idrografico");
            row.Cells[4].AddParagraph("Ubicazione");
            row.Cells[5].AddParagraph("Quota zero idrometrico (m.s.l.m.)");
            row.Cells[6].AddParagraph("S1 (m)");
            row.Cells[7].AddParagraph("S2 (m)");
            row.Cells[8].AddParagraph("S3 (m)");
            row.Cells[9].AddParagraph("Altezza idrometrica registrata h(m)");

            MigraDoc.DocumentObjectModel.Paragraph p = row.Cells[10].AddParagraph();
            p.AddText("Tendenza variazione livello\n- Aumento (");
            FormattedText ft = p.AddFormattedText("↑ ");
            ft.Color = Colors.Red;
            p.AddText("< 5% - ");
            ft = p.AddFormattedText("↑↑ ");
            ft.Color = Colors.Red;
            p.AddText("≥ 5%)  \n- Stabile (");
            ft = p.AddFormattedText("=");
            ft.Color = Colors.Blue;
            p.AddText(")  \n- Diminuzione (");
            ft = p.AddFormattedText("↓ ");
            ft.Color = Colors.Green;
            p.AddText("< 5% - ");
            ft = p.AddFormattedText("↓↓ ");
            ft.Color = Colors.Green;
            p.AddText("≥ 5%)");

            for (int i = 0; i < tableAll3.Columns.Count; i++)
            {
                row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
            }

            row.Cells[10].Format.Alignment = ParagraphAlignment.Left;
        }

        private void ContentTableAll3()
        {
            string riga = fileCSV3.ReadLine();
            string[] elencoCampi = { "" };

            // ****** Prime due letture del file per ottenere numero di bacini e numero di stazioni per ogni bacino ******

            int numBacini = 0;
            int indice = -1;
            int[] numStazioniPerBacino;
            int[] numRighePerBacino; // Array che conterrà a ogni posizione il numero di righe da inserire nel PDF per quel bacino
            double s1 = 0, s2 = 0, s3 = 0, variazione = 0;
            double altezzaRegistrata = -1;

            // Prima lettura file: vedo quanti sono i bacini
            while (riga != null)
            {
                elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                if ((elencoCampi[0] != "") && (elencoCampi[1] == ""))
                {
                    // La riga appena letta è l'intestazione di un bacino perchè la prima cella non è vuota
                    // (c'è scritto il nome del bacino), ma quella dopo sì (è sufficiente come controllo?)
                    numBacini++;
                }

                riga = fileCSV3.ReadLine();
            }

            if (numBacini == 0)
            {
                // Il file CSV relativo all'allegato3 era vuoto, lo segnalo nel log
                log.WriteToLog("Allegato A - File allegato3.csv vuoto!", level.Warning);
                Row row = tableAll3.AddRow();
                row.Format.Font.Bold = false;
                row.Cells[0].AddParagraph("Dati non presenti");
                row.Cells[0].MergeRight = 10;
                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            }
            else
            {
                // Scrivo nel file PDF solo se c'erano dati nel CSV
                numStazioniPerBacino = new int[numBacini]; // Array che conterrà a ogni posizione il numero di stazioni per quel bacino
                numRighePerBacino = new int[numBacini]; 
                fileCSV3.BaseStream.Seek(0, SeekOrigin.Begin); // Torno a inizio file csv

                riga = fileCSV3.ReadLine();
                elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                // Seconda lettura file: quante stazioni per ogni bacino
                while (riga != null)
                {
                    elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                    if ((elencoCampi[0] != "") && (elencoCampi[1] == "")) indice++; // La riga è l'intestazione di un bacino

                    if ((elencoCampi[0] != "") && (elencoCampi[1] != ""))
                    {
                        // La riga appena letta è l'intestazione di una stazione perchè i campi non sono vuoti
                        numStazioniPerBacino[indice]++;

                        // Controllo se la riga contiene dati che superano la soglia
                        if ((elencoCampi[6] != "") && (elencoCampi[9] != ""))
                        {
                            altezzaRegistrata = Convert.ToDouble(elencoCampi[9]);
                            s1 = Convert.ToDouble(elencoCampi[6]);
                            if (altezzaRegistrata > s1)
                            {
                                numRighePerBacino[indice]++;
                            }
                        }
                    }

                    riga = fileCSV3.ReadLine();
                }

                // ****** A questo punto costruzione del contenuto del PDF con le informazioni appena recuperate ******

                fileCSV3.BaseStream.Seek(0, SeekOrigin.Begin); // Torno a inizio file csv

                for (int b = 0; b < numBacini; b++)
                {
                    if (numRighePerBacino[b] > 0)
                    {
                        // All'indice i dell'array numStazioniPerBacino ci sarà il numero di stazioni per l'i-esimo bacino
                        riga = fileCSV3.ReadLine(); // La prima riga dovrebbe contenere solo il nome del bacino, lo riporto nel PDF
                        // Elimino il primo carattere del CSV, se è il char di controllo della codifica 
                        if ((b == 0) && (riga.ToCharArray()[0] == 65279))
                            riga = riga.Remove(0, 1);

                        elencoCampi = CSVRowToStringArray(riga, ';', '\n');
                        Row row = tableAll3.AddRow();
                        row.Format.Font.Bold = false;
                        row.Cells[0].Shading.Color = Colors.LawnGreen;
                        row.Cells[0].AddParagraph(elencoCampi[0]);
                        row.Cells[0].MergeRight = 10;
                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;

                        // Poi ci sono le righe con dati di ogni stazione
                        for (int s = 0; s < numStazioniPerBacino[b]; s++)
                        {
                            riga = fileCSV3.ReadLine();
                            elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                            if ((elencoCampi[6] != "") && (elencoCampi[9] != ""))
                            {
                                altezzaRegistrata = Convert.ToDouble(elencoCampi[9]);
                                s1 = Convert.ToDouble(elencoCampi[6]);
                                if (altezzaRegistrata > s1)
                                {
                                    // Scrivo la riga nel PDF solo se l'altezza registrata supera almeno la soglia S1
                                    row = tableAll3.AddRow();

                                    for (int i = 0; i < (elencoCampi.Length - 1); i++)
                                    {
                                        // Riempio tutta la riga con i valori del CSV, tranne l'ultima cella perchè 
                                        // vanno inseriti dei simboli a seconda dell'entità della variazione
                                        row.Cells[i].AddParagraph(elencoCampi[i]);

                                        row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                                        if (i == 0) row.Cells[0].Format.Alignment = ParagraphAlignment.Left;

                                        if (i == 9)
                                        {
                                            if (elencoCampi[7] != "") s2 = Convert.ToDouble(elencoCampi[7]);
                                            else s2 = 0;
                                            if (elencoCampi[8] != "") s3 = Convert.ToDouble(elencoCampi[8]);
                                            else s3 = 0;
                                            if (elencoCampi[10] != "") variazione = Convert.ToDouble(elencoCampi[10]);
                                            else variazione = 0;

                                            if ((altezzaRegistrata <= s1) && (altezzaRegistrata >= 0))
                                                row.Cells[9].Shading.Color = Colors.Green;
                                            else if ((altezzaRegistrata > s1) && (altezzaRegistrata < s2) && (altezzaRegistrata >= 0))
                                                row.Cells[9].Shading.Color = Colors.Yellow;
                                            else if ((altezzaRegistrata >= s2) && (altezzaRegistrata <= s3) && (altezzaRegistrata >= 0))
                                                row.Cells[9].Shading.Color = Colors.Orange;
                                            else if ((altezzaRegistrata >= s3) && (altezzaRegistrata >= 0))
                                            {
                                                row.Cells[9].Shading.Color = Colors.Red;
                                                row.Cells[9].Format.Font.Color = Colors.White;
                                            }

                                            row.Cells[10].Format.Font.Bold = true;

                                            if (variazione == 0)
                                            {
                                                row.Cells[10].AddParagraph("=");
                                                row.Cells[10].Format.Font.Color = Colors.Blue;
                                            }
                                            else if ((variazione < 5) && (variazione > 0))
                                            {
                                                row.Cells[10].AddParagraph("↑");
                                                row.Cells[10].Format.Font.Color = Colors.Red;
                                            }
                                            else if (variazione >= 5)
                                            {
                                                row.Cells[10].AddParagraph("↑↑");
                                                row.Cells[10].Format.Font.Color = Colors.Red;
                                            }
                                            else if ((variazione > -5) && (variazione < 0))
                                            {
                                                row.Cells[10].AddParagraph("↓");
                                                row.Cells[10].Format.Font.Color = Colors.Green;
                                            }
                                            else if (variazione <= 5)
                                            {
                                                row.Cells[10].AddParagraph("↓↓");
                                                row.Cells[10].Format.Font.Color = Colors.Green;
                                            }

                                            row.Cells[10].VerticalAlignment = VerticalAlignment.Center;
                                            row.Cells[10].Format.Font.Bold = true;
                                            row.Cells[5].Format.Alignment = row.Cells[6].Format.Alignment = row.Cells[7].Format.Alignment = 
                                                row.Cells[8].Format.Alignment = row.Cells[9].Format.Alignment = ParagraphAlignment.Center;

                                        } // if (i == 9)

                                    }// for (scorre gli elementi di una riga)
                                } // if (soglia superata)

                            } // if (altezza registrata e s1 non sono vuoti)

                        } // for (scorre le stazioni di un bacino)
                    } // if (ci sono righe da scrivere per quel bacino)

                    else
                    {
                        // Se per un certo bacino non ho inserito righe nel PDF, devo solo saltare tutte le righe del CSV relative a quel bacino
                        // num righe da saltare = 1 con il nome del bacino + quante stazioni ci sono nel bacino
                        for (int salta = 0; salta <= numStazioniPerBacino[b]; salta++)
                        {
                            riga = fileCSV3.ReadLine();
                            elencoCampi = CSVRowToStringArray(riga, ';', '\n');
                        }
                    }

                } // for (scorre i bacini)

                // Se nessuna stazione superava la soglia, o se il file CSV era vuoto, aggiungo una riga alla tabella in cui lo scrivo
                int sommaRighe = 0;
                for (int b = 0; b < numBacini; b++)
                    sommaRighe = sommaRighe + numRighePerBacino[b];
                if (sommaRighe == 0)
                {
                    Row row = tableAll3.AddRow();
                    row.Cells[0].MergeRight = 10;
                    row.Cells[0].Format.Font.Bold = true;
                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    if (numBacini == 0)
                        row.Cells[0].AddParagraph("Dati non presenti.");
                    else row.Cells[0].AddParagraph("Nessuna stazione supera le soglie. ");
                }

            } //else  
        }

        private void LegendaAll3()
        {
            //Creo la legenda dove vengono indicati i criteri usati nella colorazione delle celle, da mettere dopo la fine della tabella con i dati
            MigraDoc.DocumentObjectModel.Paragraph nota = new MigraDoc.DocumentObjectModel.Paragraph();
            nota.AddFormattedText("Per la definizione delle soglie e per ulteriori informazioni verificare le monografie pubblicate nell\'apposita sezione del sito Internet della " +
                "Protezione Civile al link: http://www.sardegnaambiente.it/protezionecivile/nowcasting/monografie_idrometri.html ");
            nota.Format.Font.Italic = true;
            nota.Format.Alignment = ParagraphAlignment.Left;
            nota.Format.Font.Size = 7;

            //Row row = tableAll3.AddRow(); 
            //row.Height = 4;
            //row.Borders.Visible = false;
            //row.Shading.Color = Colors.White;
            
            //int altezza = 12;
            //row = tableAll3.AddRow();
            //row.Height = altezza;
            //row.Format.Font.Size = 9;
            //row.Cells[0].MergeRight = 7;
            //row.Cells[0].AddParagraph("Legenda dei colori");
            //row.Cells[0].Shading.Color = grigio;
            //row.Cells[0].Format.Font.Bold = true;
            //row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            //row.Cells[0].Format.Font.Color = Colors.White;
            //row.Cells[0].Borders.Bottom.Visible = false;
            //row.Cells[8].Borders.Visible = row.Cells[9].Borders.Visible = row.Cells[10].Borders.Visible = false;
            //row.Cells[8].Shading.Color = Colors.White;

            //row = tableAll3.AddRow();
            //row.Height = altezza;
            //row.Cells[0].MergeRight = 1;
            //row.Cells[0].Shading.Color = Colors.LimeGreen;
            //row.Cells[0].AddParagraph("h ≤ S1");
            //row.Cells[2].MergeRight = 5;
            //row.Cells[0].Borders.Bottom.Visible = row.Cells[2].Borders.Bottom.Visible = false;
            //row.Cells[0].Borders.Top.Visible = row.Cells[2].Borders.Top.Visible = false;
            //row.Cells[8].Borders.Visible = row.Cells[9].Borders.Visible = row.Cells[10].Borders.Visible = false;
            //row.Cells[8].Shading.Color = Colors.White;
            //row.Cells[2].AddParagraph("Livello idrometrico inferiore alla PRIMA soglia");
            //row.Cells[2].Format.Alignment = ParagraphAlignment.Left;

            //row = tableAll3.AddRow();
            //row.Height = altezza;
            //row.Cells[0].MergeRight = 1;
            //row.Cells[0].Shading.Color = Colors.Yellow;
            //row.Cells[0].AddParagraph("S1 < h < S2");
            //row.Cells[2].MergeRight = 5;
            //row.Cells[0].Borders.Bottom.Visible = row.Cells[2].Borders.Bottom.Visible = false;
            //row.Cells[0].Borders.Top.Visible = row.Cells[2].Borders.Top.Visible = false;
            //row.Cells[8].Borders.Visible = row.Cells[9].Borders.Visible = row.Cells[10].Borders.Visible = false;
            //row.Cells[8].Shading.Color = Colors.White;
            //row.Cells[2].AddParagraph("Livello idrometrico compreso tra la PRIMA e la SECONDA soglia");
            //row.Cells[2].Format.Alignment = ParagraphAlignment.Left;

            //row = tableAll3.AddRow();
            //row.Height = altezza;
            //row.Cells[0].MergeRight = 1;
            //row.Cells[0].Shading.Color = Colors.Orange;
            //row.Cells[0].AddParagraph("S2 ≤ h < S3");
            //row.Cells[2].MergeRight = 5;
            //row.Cells[0].Borders.Bottom.Visible = row.Cells[2].Borders.Bottom.Visible = false;
            //row.Cells[0].Borders.Top.Visible = row.Cells[2].Borders.Top.Visible = false;
            //row.Cells[8].Borders.Visible = row.Cells[9].Borders.Visible = row.Cells[10].Borders.Visible = false;
            //row.Cells[8].Shading.Color = Colors.White;
            //row.Cells[2].AddParagraph("Livello idrometrico compreso tra la SECONDA e la TERZA soglia");
            //row.Cells[2].Format.Alignment = ParagraphAlignment.Left;

            //row = tableAll3.AddRow();
            //row.Height = altezza;
            //row.Cells[0].MergeRight = 1;
            //row.Cells[0].Shading.Color = Colors.Red;
            //row.Cells[0].AddParagraph("h ≥ S3");
            //row.Cells[2].MergeRight = 5;
            //row.Cells[0].Borders.Top.Visible = row.Cells[2].Borders.Top.Visible = false;
            //row.Cells[8].Borders.Visible = row.Cells[9].Borders.Visible = row.Cells[10].Borders.Visible = false;
            //row.Cells[8].Shading.Color = Colors.White;
            //row.Cells[2].AddParagraph("Livello idrometrico superiore alla TERZA soglia");
            //row.Cells[2].Format.Alignment = ParagraphAlignment.Left;

            //row = tableAll3.AddRow(); 
            //row.HeightRule = RowHeightRule.Auto;
            //row.Cells[0].MergeRight = 7;
            //row.Cells[0].Borders.Visible = false; 
            //row.Cells[0].Add(nota);
            //row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            //row.Cells[8].Borders.Visible = row.Cells[9].Borders.Visible = row.Cells[10].Borders.Visible = false;
            //row.Cells[8].Shading.Color = Colors.White;

            section.AddParagraph().AddLineBreak(); 

            Table legenda = section.AddTable();
            legenda.Borders.Width = 0.2;
            legenda.Rows.LeftIndent = 0;
            legenda.Format.Alignment = ParagraphAlignment.Center;
            legenda.Rows.VerticalAlignment = VerticalAlignment.Center;
            legenda.Format.Font.Size = 6;

            legenda.AddColumn("4.5cm");
            legenda.AddColumn("8.5cm");
            legenda.KeepTogether = true;

            int altezza = 12;
            Row firstrow = legenda.AddRow();
            Row row = firstrow;
            row.Height = altezza;
            row.Format.Font.Size = 9;
            row.Cells[0].MergeRight = 1;
            row.Cells[0].AddParagraph("Legenda dei colori");
            row.Cells[0].Shading.Color = grigio;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].Format.Font.Color = Colors.White;
            row.Cells[0].Borders.Bottom.Visible = false;

            row = legenda.AddRow();
            row.KeepWith = firstrow.Index;
            row.Height = altezza;
            row.Cells[0].Shading.Color = Colors.LimeGreen;
            row.Cells[0].AddParagraph("h ≤ S1");
            row.Cells[0].Borders.Bottom.Visible = row.Cells[1].Borders.Bottom.Visible = false;
            row.Cells[0].Borders.Top.Visible = row.Cells[1].Borders.Top.Visible = false;
            row.Cells[1].AddParagraph("Livello idrometrico inferiore alla PRIMA soglia");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;

            row = legenda.AddRow();
            row.KeepWith = firstrow.Index;
            row.Height = altezza;
            row.Cells[0].Shading.Color = Colors.Yellow;
            row.Cells[0].AddParagraph("S1 < h < S2");
            row.Cells[0].Borders.Bottom.Visible = row.Cells[1].Borders.Bottom.Visible = false;
            row.Cells[0].Borders.Top.Visible = row.Cells[1].Borders.Top.Visible = false;
            row.Cells[1].AddParagraph("Livello idrometrico compreso tra la PRIMA e la SECONDA soglia");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;

            row = legenda.AddRow();
            row.KeepWith = firstrow.Index;
            row.Height = altezza;
            row.Cells[0].Shading.Color = Colors.Orange;
            row.Cells[0].AddParagraph("S2 ≤ h < S3");
            row.Cells[0].Borders.Bottom.Visible = row.Cells[1].Borders.Bottom.Visible = false;
            row.Cells[0].Borders.Top.Visible = row.Cells[1].Borders.Top.Visible = false;
            row.Cells[1].AddParagraph("Livello idrometrico compreso tra la SECONDA e la TERZA soglia");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;

            row = legenda.AddRow();
            row.KeepWith = firstrow.Index;
            row.Height = altezza;
            row.Cells[0].Shading.Color = Colors.Red;
            row.Cells[0].Format.Font.Color = Colors.White;
            row.Cells[0].AddParagraph("h ≥ S3");
            row.Cells[0].Borders.Top.Visible = row.Cells[1].Borders.Top.Visible = false;
            row.Cells[1].AddParagraph("Livello idrometrico superiore alla TERZA soglia");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;

            //row = legenda.AddRow();
            //row.KeepWith = firstrow.Index;
            //row.HeightRule = RowHeightRule.Auto;
            //row.Cells[0].MergeRight = 1;
            //row.Cells[0].Borders.Visible = false;
            //row.Cells[0].Add(nota);
            //row.Cells[0].Format.Alignment = ParagraphAlignment.Left;

            firstrow.KeepWith = legenda.Rows.Count - 1;
            section.Add(nota);
    
        }

        void CampiCompilabili()
        {
            section.AddParagraph().AddLineBreak();

            int altezzaRigheTitoli = 15;
            int altezzaRigheTesto = 110;

            DocumentRenderer docRenderer = new DocumentRenderer(document); 
            int pageCount = 0;
   
            Table t1 = section.AddTable();
            t1.Borders.Width = 0.2;
            t1.Rows.LeftIndent = 0;
            t1.Format.Alignment = ParagraphAlignment.Left;
            t1.Rows.VerticalAlignment = VerticalAlignment.Center;
            t1.Format.Font.Size = 9;
            t1.AddColumn("20cm");
            t1.KeepTogether = true;

            Row row = t1.AddRow();
            TextFrame tf = row.Cells[0].AddTextFrame();
            row.Format.Font.Bold = true;
            row.Borders.Bottom.Visible = false;
            row.Height = altezzaRigheTitoli;

            row = t1.AddRow();
            row.Height = altezzaRigheTesto;
            row.Borders.Top.Visible = false;

            // Le due righe seguenti mi danno in pageCount l'informazione aggiornata su a che pagina siamo arrivati
            docRenderer.PrepareDocument();
            pageCount = docRenderer.FormattedDocument.PageCount;
            DisegnaCampoCompilabile(615, pageCount);

            row = t1.AddRow();
            row.Height = altezzaRigheTitoli;
            row.Borders.Visible = false;

            Table t2 = section.AddTable();
            t2.Borders.Width = 0.2;
            t2.Rows.LeftIndent = 0;
            t2.Format.Alignment = ParagraphAlignment.Left;
            t2.Rows.VerticalAlignment = VerticalAlignment.Center;
            t2.Format.Font.Size = 9;
            t2.AddColumn("20cm");
            t2.KeepTogether = true;

            row = t2.AddRow();
            row.Cells[0].AddParagraph("Valutazione meteorologica");
            row.Format.Font.Bold = true;
            row.Shading.Color = Colors.Gray;
            row.Format.Font.Color = Colors.White;
            row.Height = altezzaRigheTitoli;
            
            row = t2.AddRow();
            row.Height = altezzaRigheTesto;

            row = t2.AddRow();
            row.Shading.Color = Colors.Gray;
            row.Height = altezzaRigheTitoli;

            row = t2.AddRow();
            row.Height = altezzaRigheTitoli;
            row.Borders.Visible = false;

            Table t3 = section.AddTable();
            t3.Borders.Width = 0.2;
            t3.Rows.LeftIndent = 0;
            t3.Format.Alignment = ParagraphAlignment.Left;
            t3.Rows.VerticalAlignment = VerticalAlignment.Center;
            t3.Format.Font.Size = 9;
            t3.AddColumn("20cm");
            t3.KeepTogether = true;

            row = t3.AddRow();
            row.Cells[0].AddParagraph("Valutazioni idrauliche");
            row.Format.Font.Bold = true;
            row.Shading.Color = Colors.Gray;
            row.Format.Font.Color = Colors.White;
            row.Height = altezzaRigheTitoli;
            
            row = t3.AddRow();
            row.Height = altezzaRigheTesto;

            row = t3.AddRow();
            row.Shading.Color = Colors.Gray;
            row.Height = altezzaRigheTitoli;

            row = t3.AddRow();
            row.Height = altezzaRigheTitoli * 2;
            row.Borders.Visible = false;

            Table t4 = section.AddTable();
            t4.Format.Alignment = ParagraphAlignment.Center;
            t4.Borders.Visible = false;
            t4.AddColumn("20cm");
            t4.KeepTogether = true;
            
            row = t4.AddRow();
            row.Format.Font.Size = 9;
            row.Cells[0].AddParagraph("Il Direttore del Servizio Previsione rischi e dei sistemi informativi, infrastrutture e reti");
            
            row = t4.AddRow();
            row.Height = altezzaRigheTitoli * 2;
            row.Format.Font.Size = 9;

            docRenderer.PrepareDocument();
            pageCount = docRenderer.FormattedDocument.PageCount;

           
        }

        void DisegnaCampoCompilabile(int yll, int pag, bool last=false)
        {
            int xll = 15;
            int width = 565;
            //int yll = 615; // QUESTA è la vera incognita, insieme al numero di pagina 
            int height = 105;

            //variables
            String pathout = "prova.pdf"; //TODO: Se last = true, salverò tutto in AllegatoA.pdf, se no in un tmp.pdf  

            iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(filename);
            //Seleziono tutte le pagine del documento appena creato
            int numPages = reader.NumberOfPages;
            string range = "1-";
            range += numPages.ToString();
            reader.SelectPages(range);

            //L'oggetto PdfStamper acquisisce le pagine del pdf e mi permette di modificarle 
            PdfStamper stamper = new PdfStamper(reader, new FileStream(pathout, FileMode.Create));
            // PdfContentByte from stamper to add content to the pages over the original content

            PdfContentByte pbover = stamper.GetOverContent(pag);
            PdfWriter writer = stamper.Writer;

            TextField p = new TextField(writer, new Rectangle(xll, yll, (xll + width), (yll + height)), "nome");
            p.Text = "Commento"; // DA TOGLIERE 
            p.Alignment = Element.ALIGN_TOP;
            p.Options = PdfFormField.FF_MULTILINE;
            p.TextColor = BaseColor.BLACK;
            p.BorderWidth = 0;
            p.BackgroundColor = BaseColor.RED; // Da mettere a bianco
            p.BorderColor = BaseColor.WHITE;
            p.FontSize = 9;

            p.GetTextField().SetFieldFlags(PdfFormField.FF_MULTILINE);

            stamper.AddAnnotation(p.GetTextField(), pag);
            stamper.Close();
        }


        void DefineStyles()
        {
            // Da rivedere se c'è qualcosa da modificare negli stili
            Style style = document.Styles["Normal"];
            style.Font.Name = "Arial";
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.Font.Size = 8;

        } //DefineStyles()

        // CSVRowToStringArray: Funzione che prende una riga da un csv e torna un array di stringhe con l'elenco dei contenuti di ogni elemento
        // Parametri: stringa con riga da esaminare, carattere che separa i campi, carattere che separa le righe
        private static string[] CSVRowToStringArray(string r, char fieldSep, char stringSep)
        {
            if (r != null)
            {
                bool bolQuote = false;
                StringBuilder str = new StringBuilder();
                List<string> ret = new List<string>();

                foreach (char c in r.ToCharArray())
                    if ((c == fieldSep && !bolQuote))
                    {
                        ret.Add(str.ToString());
                        str.Clear();
                    }
                    else
                        if (c == stringSep)
                            bolQuote = !bolQuote;
                        else
                            str.Append(c);

                ret.Add(str.ToString());
                return ret.ToArray();
            }
            else return null;
        }

    } //class AllA
} // namespace
