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

namespace BollettiniMonitoraggio
{
    class All1
    {
        protected Document document;
        protected Table table;
        protected Section section;
        protected System.IO.StreamReader fileCSV;
        protected string nomeFileCSV; // Nome del CSV da aprire, passato come parametro
        protected string consultazione; // stringa che contiene ora e data della creazione del file - messa nel costruttore per impedire che possa scattare un minuto tra una pagina e l'altra 
        protected string dirigente; // Nome del dirigente che firma il PDF, da prelevare da config.ini

        readonly static Color azzurro1 = new Color(171, 205, 239);
        readonly static Color azzurro2 = new Color(0, 127, 255);
        readonly static Color azzurro3 = new Color(0, 0, 128);
        readonly static Color grigio = new Color(215, 215, 215);
        readonly static Color verde = Colors.LawnGreen;
        readonly static Color giallo = Colors.LightYellow;

        private LogWriter log; // collegamento al file di log dove scriverò informazioni sull'esecuzione
        protected string filename; // nome file pdf

        public All1(string n)
        {
            consultazione = "Estrazione dati delle ore " + DateTime.Now.ToString("HH:mm") + " del " + DateTime.Now.ToString("dd/MM/yyyy");
            log = LogWriter.Instance; // File di log in cui scrivo informazioni su esecuzione ed errori
            log.WriteToLog("ALL. 1 - Inizio scrittura bollettino, " + consultazione + ".", level.Info);
            dirigente = "";

            try
            {
                // Imposta il file INI in cui andrò a leggere il nome e cognome del dirigente responsabile,
                // che salvo in una stringa per scriverlo poi nel footer del PDF
                IniParser fileIni = new IniParser(@".\BollettinoMonitoraggio.ini");
                string tmp = fileIni.GetSetting("SETTING_SECTION", "responsabile");
                dirigente = tmp.Replace('_', ' ');
                tmp = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToLower(dirigente);
                dirigente = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(tmp);
            }
            catch (System.Exception err)
            {
                // Intercetta la System.IO.FileNotFoundException nel caso non sia stato trovato il file .ini
                err.Source = "IniParser";
                log.WriteToLog("ALL. 1 - BollettinoPiogge: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception);
            }

            filename = "Allegato1.pdf";
            nomeFileCSV = n;
            CreaPDFdaCSV();
            log.WriteToLog("ALL. 1 - Fine scrittura bollettino, " + consultazione + ".", level.Info);
        }

        private void CreaPDFdaCSV()
        {
            try
            {
                using (fileCSV = new StreamReader(nomeFileCSV, true))
                {
                    // Creazione del PDF
                    document = new Document();
                    document.Info.Title = "Bollettino di analisi piogge";
                    document.Info.Author = "R.A.S.";

                    document.UseCmykColor = true;
                    const bool unicode = true;
                    const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

                    DefineStyles();
                    // Impaginazione delle informazioni nel pdf
                    CreatePage();

                    PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode, embedding); // opzioni di visualizzazione
                    pdfRenderer.Document = document;
                    pdfRenderer.RenderDocument();

                    // Alla fine salva il pdf appena creato: se esiste già con lo stesso nome lo elimina e poi salva quello nuovo
                    if (System.IO.File.Exists(filename))
                    {
                        FileInfo bInfoOld = new FileInfo(filename);
                        bInfoOld.IsReadOnly = false;
                        System.IO.File.Delete(filename);
                    }

                    pdfRenderer.PdfDocument.Save(filename);

                    // Imposto il file come accessibile in sola lettura
                    FileInfo fInfo = new FileInfo(filename);
                    fInfo.IsReadOnly = true;

                    // System.Diagnostics.Process.Start(filename);
                } //using
            } //try

            catch (System.Exception err)
            {
                // Intercetta tutti i tipi di eccezione ma quelle che si dovrebbero verificare più di frequente sono:
                // System.IO.FileNotFoundException, System.IO.DirectoryNotFoundException, System.IO.IOException...        
                err.Source = "CreaPDFdaCSV()";
                // Salvo nel log il messaggio di errore con un pò di informazioni sulla funzione che ha lanciato l'eccezione e sul tipo di eccezione
                log.WriteToLog("ALL. 1 - BollettinoPiogge: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception);
            }
        }


        void CreatePage()
        {
            // Each MigraDoc document needs at least one section.
            section = document.AddSection();
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Portrait;
            section.PageSetup.TopMargin = "05mm";
            section.PageSetup.LeftMargin = "05mm";
            section.PageSetup.RightMargin = "05mm";
            section.PageSetup.BottomMargin = "15mm";
            section.PageSetup.DifferentFirstPageHeaderFooter = false;

            section.PageSetup.FooterDistance = Unit.FromCentimeter(0.2);

            // Aggiungo i numeri di pagina
            section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            Paragraph paragraph = new Paragraph();
            paragraph.Format.Font.Size = 7;
            paragraph.Format.Alignment = ParagraphAlignment.Right;
            paragraph.AddText("Pagina ");
            paragraph.AddPageField();
            paragraph.AddText(" di ");
            paragraph.AddNumPagesField();
            section.Footers.Primary.Add(paragraph);
            section.Footers.EvenPage.Add(paragraph.Clone());

            // In fondo al documento aggiungo la firma del dirigente
            Paragraph firma = new Paragraph();
            firma.AddFormattedText("ARPAS\nF.to il Dirigente Responsabile\n" + dirigente);
            firma.Format.LeftIndent = 0;
            firma.Format.RightIndent = Unit.FromCentimeter(15);
            firma.Format.Alignment = ParagraphAlignment.Center;
            firma.Format.Font.Size = 7;
            firma.Format.Font.Bold = true;
            section.Footers.Primary.Add(firma);
            section.Footers.EvenPage.Add(firma.Clone());

            // ******************* Da qui in poi - creazione tabella *******************

            CreateTableHeader(); // header (viene replicato in ogni pagina)

            CreateTableBody(); // Dati bollettino

            Legenda(); // Legenda sulla colorazione delle celle in caso di superamento soglie

            // Inserisco il disclaimer sui dati 
            Paragraph disclaimer = document.LastSection.AddParagraph();
            disclaimer.Format.Font.Bold = false;
            disclaimer.Format.Font.Size = 6;
            disclaimer.Format.Alignment = ParagraphAlignment.Center;
            FormattedText fd = disclaimer.AddFormattedText("\n\"Composizione e rappresentazione dei dati eseguita con modalità automatiche su dati della rete di stazioni meteorologiche fiduciarie della Regione Sardegna gestita dall\'Agenzia per la Protezione dell'Ambiente della Sardegna, ARPAS, acquisiti in tempo reale e sottoposti ad un processo automatico di validazione di primo livello\"");

        } //CreatePage()        


        void CreateTableHeader()
        {
            // Crea l'intestazione della tabella che andrà ripetuta in ogni pagina
            table = section.AddTable();
            table.Borders.Width = 0.2;
            table.Rows.LeftIndent = 0;
            table.Format.Alignment = ParagraphAlignment.Center;
            const int numColonne = 13;
            const int altezzaRighe = 12;
            table.Format.Font.Size = 6;
            table.TopPadding = 1;
            table.BottomPadding = 1;

            // Definisco le colonne
            Column column = table.AddColumn("0.6cm");
            column = table.AddColumn("1.9cm");
            column = table.AddColumn("2cm");
            column = table.AddColumn("2.5cm");
            column = table.AddColumn("1cm");
            column = table.AddColumn("1cm");
            column.Shading.Color = giallo;
            column = table.AddColumn("1cm");
            column.Shading.Color = giallo;
            column = table.AddColumn("1.5cm");
            column.Shading.Color = giallo;
            column = table.AddColumn("2.5cm");
            column.Shading.Color = giallo;
            column = table.AddColumn("2cm");
            column = table.AddColumn("1cm");
            column = table.AddColumn("1.5cm");
            column = table.AddColumn("1.5cm");
            table.Rows.Height = altezzaRighe;

            //// Logo Regione
            //Row row = table.AddRow();
            //row.HeadingFormat = true;
            //row.Borders.Visible = false;
            //Image image = new Image("./logo.bmp");
            //image.Height = "2.3cm";
            //image.Width = "4cm";
            //row.Cells[0].MergeRight = 5;
            //image.Left = ShapePosition.Center;
            //row.Cells[6].MergeRight = 6;
            //row.Cells[6].Add(image);
            //row.Shading.Color = Colors.White;

            //// Intestazione
            //row = table.AddRow();
            //row.HeadingFormat = true;
            //row.Cells[0].MergeRight = 12;
            //row.Cells[0].AddParagraph("\nDirezione Generale della Protezione Civile \nCentro Funzionale Decentrato");
            //row.Format.Font.Size = 8;
            //row.Format.Alignment = ParagraphAlignment.Left;
            //row.Borders.Visible = false;

            // Logo Regione
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Borders.Visible = false;
            row.Cells[0].MergeRight = 12;
            MigraDoc.DocumentObjectModel.Shapes.Image image = new MigraDoc.DocumentObjectModel.Shapes.Image("./logo RAS con didascalia.png");
            image.LockAspectRatio = true;
            image.Height = Unit.FromCentimeter(2);
            image.Width = Unit.FromCentimeter(4.5);
            row.Cells[0].Add(image);

            // Titolo
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Cells[0].MergeRight = 12;
            DateTime dateValue = DateTime.Now;
            String titolo = "\n\nANALISI DELLA PIOGGIA REGISTRATA NELLE ULTIME 24 ORE DALLE STAZIONI PLUVIOMETRICHE DELLA RETE FIDUCIARIA";
            row.Cells[0].AddParagraph(titolo);
            row.Cells[0].AddParagraph(consultazione).Format.Font.Italic = true;
            row.Format.Font.Size = 9;
            row.Format.Font.Bold = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Borders.Visible = false;
            row.Cells[0].AddParagraph();

            // Intestazione tabella
            row = table.AddRow();
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

            Paragraph p = new Paragraph();
            p.AddFormattedText("h/h");
            FormattedText ft = p.AddFormattedText("Tr20anni");
            ft.Subscript = true;
            row.Cells[11].Add(p);
            row.Cells[11].MergeDown = 1;

            Paragraph p1 = new Paragraph();
            p1.AddFormattedText("h/h");
            FormattedText ft1 = p1.AddFormattedText("Tr100anni");
            ft1.Subscript = true;
            row.Cells[12].Add(p1);
            row.Cells[12].MergeDown = 1;

            for (int i = 0; i < numColonne; i++)
            {
                row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
            }

            row = table.AddRow();
            row.Format.Font.Bold = true;
            Paragraph p2 = new Paragraph();
            p2.AddFormattedText("h");
            FormattedText ft2 = p2.AddFormattedText("Tr20anni");
            ft2.Subscript = true;
            row.Cells[5].Add(p2);
            Paragraph p3 = new Paragraph();
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

        void CreateTableBody()
        {
            const int numRighePerStazione = 5; // sappiamo che per ogni stazione ci sono sempre 5 righe
            string riga = fileCSV.ReadLine();
            string[] elencoCampi = { "" };
            bool alternaColoreRighe = true; // Mi serve per colorare una stazione di grigio e una lasciarla bianca, per facilitare la lettura

            // ****** Prime due letture del file per ottenere numero di bacini e numero di stazioni per ogni bacino ******
            int numBacini = 0;
            int indice = -1;

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

                riga = fileCSV.ReadLine();
            }

            if (numBacini == 0)
            {
                // Il file CSV era vuoto, lo segnalo nel log
                log.WriteToLog("ALL. 1 - File .CSV vuoto! ", level.Error);
                // Aggiungo una riga alla tabella in cui dico che il CSV era vuoto
                Row row = table.AddRow();
                row.Cells[0].MergeRight = 12;
                row.Cells[0].Shading.Color = Colors.White;
                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                row.Cells[0].Format.Font.Bold = true;
                row.Cells[0].AddParagraph("Dati non presenti.");
            }

            else
            {
                // File CSV non vuoto
                int[] numStazioniPerBacino = new int[numBacini]; // Array che conterrà a ogni posizione il numero di stazioni per quel bacino
                fileCSV.BaseStream.Seek(0, SeekOrigin.Begin); // Torno a inizio file csv
                riga = fileCSV.ReadLine();
                elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                // Seconda lettura file: quante stazioni per ogni bacino
                while ((riga != null) && (riga != ""))
                {
                    elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                    if ((elencoCampi[0] != "") && (elencoCampi[1] == "")) indice++; // La riga è l'intestazione di un bacino

                    if ((elencoCampi[0] != "") && (elencoCampi[1] != ""))
                    {
                        // La riga appena letta è l'intestazione di una stazione perchè i campi non sono vuoti
                        numStazioniPerBacino[indice]++;
                    }

                    riga = fileCSV.ReadLine();
                }

                // ****** A questo punto costruzione del contenuto del PDF con le informazioni appena recuperate ******

                fileCSV.BaseStream.Seek(0, SeekOrigin.Begin); // Torno a inizio file csv

                for (int b = 0; b < numBacini; b++)
                {
                    // All'indice i dell'array numStazioniPerBacino ci sarà il numero di stazioni per l'i-esimo bacino
                    riga = fileCSV.ReadLine(); // La prima riga dovrebbe contenere solo il nome del bacino, lo riporto nel PDF

                    // Elimino il primo carattere del CSV, se è il char di controllo della codifica 
                    if ((b == 0) && (riga.ToCharArray()[0] == 65279))
                        riga = riga.Remove(0, 1);

                    elencoCampi = CSVRowToStringArray(riga, ';', '\n');
                    Row row = table.AddRow();
                    row.HeadingFormat = true;
                    row.Format.Font.Bold = true;
                    row.Cells[0].Shading.Color = verde;
                    string maiuscole = elencoCampi[0].ToUpper();
                    row.Cells[0].AddParagraph(maiuscole);
                    row.Cells[0].MergeRight = 12;
                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;

                    // Poi ci sono le righe con dati di ogni stazione
                    for (int s = 0; s < numStazioniPerBacino[b]; s++)
                    {
                        riga = fileCSV.ReadLine();
                        elencoCampi = CSVRowToStringArray(riga, ';', '\n');
                        row = table.AddRow();

                        // Le prime 5 celle della nuova riga sono uniche per tutta la stazione (n., stazione, comune, zona di allerta, quota)
                        row.Cells[0].MergeDown = row.Cells[1].MergeDown = row.Cells[2].MergeDown = row.Cells[3].MergeDown = row.Cells[4].MergeDown = numRighePerStazione - 1;
                        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        row.Cells[0].AddParagraph(elencoCampi[0]);
                        if (alternaColoreRighe) row.Cells[0].Shading.Color = grigio;

                        for (int i = 2; i < (elencoCampi.Length-1); i++)
                        {
                            row.Cells[i - 1].AddParagraph(elencoCampi[i]);

                            if (i == 2)
                            {
                                // Sto scrivendo la cella con il nome della stazione, quindi devo aggiungere, oltre al nome,
                                // la data e ora dell'ultimo dato disponibile, che si trovano nell'ultimo elemento di elencoCampi
                                Paragraph p = new Paragraph();
                                p.AddFormattedText("Ultimo dato disponibile: ");

                                // aggiunta di data e ora ultimo dato disponibile, prendendo da csv stringa formattata come ddMMyyyyHHmm
                                string ultimodato = elencoCampi[elencoCampi.Length - 1];
                                ultimodato = ultimodato.Insert(2, "/");
                                ultimodato = ultimodato.Insert(5, "/");
                                ultimodato = ultimodato.Insert(10, " alle ");
                                ultimodato = ultimodato.Insert(18, ":");
                                p.AddFormattedText(ultimodato);

                                p.Format.Font.Italic = true;
                                p.Format.Font.Size = 4;
                                row.Cells[i - 1].Add(p);
                            }


                            if (alternaColoreRighe) row.Cells[i - 1].Shading.Color = grigio;
                            row.Cells[i - 1].VerticalAlignment = VerticalAlignment.Center;
                            // Ripristino lo sfondo giallo nelle colonne "Pioggia critica di riferimento" e "Finestra di osservazione"
                            row.Cells[5].Shading.Color = row.Cells[6].Shading.Color = row.Cells[7].Shading.Color = row.Cells[8].Shading.Color = giallo;

                            if (((i == 12) || (i == 13)) && (elencoCampi[i] != ""))
                            {
                                // Devo colorare le caselle solo se nei campi ci sono dei valori, altrimenti niente
                                if ((Convert.ToDouble(elencoCampi[i]) >= 0.5) && (Convert.ToDouble(elencoCampi[i]) < 0.75))
                                    row.Cells[i - 1].Shading.Color = azzurro1;
                                else if ((Convert.ToDouble(elencoCampi[i]) >= 0.75) && (Convert.ToDouble(elencoCampi[i]) < 1))
                                {
                                    row.Cells[i - 1].Shading.Color = azzurro2;
                                    row.Cells[i - 1].Format.Font.Color = Colors.White;
                                }
                                else if (Convert.ToDouble(elencoCampi[i]) >= 1)
                                {
                                    row.Cells[i - 1].Shading.Color = azzurro3;
                                    row.Cells[i - 1].Format.Font.Color = Colors.White;
                                }
                                else
                                    row.Cells[i - 1].Shading.Color = Colors.White;
                            }
                        }
                        // Riempio le altre 4 righe che riguardano la stazione in esame
                        for (int j = 0; j < (numRighePerStazione - 1); j++)
                        {
                            row = table.AddRow();
                            riga = fileCSV.ReadLine();
                            elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                            // Salto la colonna 1 perchè c'è scritto il bacino e non mi serve ripeterlo
                            for (int i = 2; i < (elencoCampi.Length - 1); i++)
                            {
                                row.Cells[i - 1].AddParagraph(elencoCampi[i]);
                                row.Cells[i - 1].VerticalAlignment = VerticalAlignment.Center;
                                if (alternaColoreRighe) row.Cells[i - 1].Shading.Color = grigio;

                                if (((i == 12) || (i == 13)) && (elencoCampi[i] != ""))
                                {
                                    // Devo colorare le caselle solo se nei campi ci sono dei valori, altrimenti niente
                                    if ((Convert.ToDouble(elencoCampi[i]) >= 0.5) && (Convert.ToDouble(elencoCampi[i]) < 0.75))
                                        row.Cells[i - 1].Shading.Color = azzurro1;
                                    else if ((Convert.ToDouble(elencoCampi[i]) >= 0.75) && (Convert.ToDouble(elencoCampi[i]) < 1))
                                    {
                                        row.Cells[i - 1].Shading.Color = azzurro2;
                                        row.Cells[i - 1].Format.Font.Color = Colors.White;
                                    }
                                    else if (Convert.ToDouble(elencoCampi[i]) >= 1)
                                    {
                                        row.Cells[i - 1].Shading.Color = azzurro3;
                                        row.Cells[i - 1].Format.Font.Color = Colors.White;
                                    }
                                    else
                                        row.Cells[i - 1].Shading.Color = Colors.White;
                                }
                            }
                            // Ripristino lo sfondo giallo nelle colonne "Pioggia critica di riferimento" e "Finestra di osservazione"
                            row.Cells[5].Shading.Color = row.Cells[6].Shading.Color = row.Cells[7].Shading.Color = row.Cells[8].Shading.Color = giallo;
                        }
                        alternaColoreRighe = !alternaColoreRighe;
                    }
                }
            }

        } // CreateTableBody()


        void Legenda()
        {
            //Creo la legenda dove vengono indicati i criteri usati nella colorazione delle celle, da mettere a fine file 
            const double borderWidth = 0.25;
            const String leftIndent = "15.5cm";
 
            Paragraph spazio = section.AddParagraph();
            spazio.AddLineBreak(); // Lascio una riga di spazio fra i dati e la legenda
            Paragraph legenda = document.LastSection.AddParagraph();
            legenda.AddText("LEGENDA\n");
            //legenda.Format.Borders.Width = borderWidth;
            //legenda.Format.Borders.Color = Colors.Black;
            legenda.Format.LeftIndent = leftIndent;
            legenda.Format.LineSpacingRule = LineSpacingRule.Single;
            legenda.Format.Font.Size = 6;

            Paragraph legenda2 = document.LastSection.AddParagraph();
            FormattedText ft = legenda2.AddFormattedText("h/h");
            ft = legenda2.AddFormattedText("tr");
            ft.Subscript = true;
            legenda2.AddText(" < 50%\n");
            legenda2.Format.Borders.Width = borderWidth;
            legenda2.Format.Borders.Color = Colors.Black;
            legenda2.Format.LeftIndent = leftIndent;
            legenda2.Format.Font.Size = 7;

            Paragraph legenda3 = document.LastSection.AddParagraph();
            FormattedText ft3 = legenda3.AddFormattedText("50% ≤ h/h");
            ft3 = legenda3.AddFormattedText("tr");
            ft3.Subscript = true;
            legenda3.AddText(" < 75%\n");
            legenda3.Format.Borders.Width = borderWidth;
            legenda3.Format.Borders.Color = Colors.Black;
            legenda3.Format.LeftIndent = leftIndent;
            legenda3.Format.Shading.Color = azzurro1;
            legenda3.Format.Font.Size = 7;

            Paragraph legenda4 = document.LastSection.AddParagraph();
            FormattedText ft4 = legenda4.AddFormattedText("75% ≤ h/h");
            ft4 = legenda4.AddFormattedText("tr");
            ft4.Subscript = true;
            legenda4.AddText(" < 100%\n");
            legenda4.Format.Borders.Width = borderWidth;
            legenda4.Format.Borders.Color = Colors.Black;
            legenda4.Format.LeftIndent = leftIndent;
            legenda4.Format.Shading.Color = azzurro2;
            legenda4.Format.Font.Size = 7;

            Paragraph legenda5 = document.LastSection.AddParagraph();
            legenda5.Format.Font.Size = 7;
            FormattedText ft5 = legenda5.AddFormattedText("h/h");
            ft5.Color = Colors.White;
            ft5 = legenda5.AddFormattedText("tr");
            ft5.Color = Colors.White;
            ft5.Subscript = true;
            ft5 = legenda5.AddFormattedText(" ≥ 100%\n");
            ft5.Color = Colors.White;
            legenda5.Format.Borders.Width = borderWidth;
            legenda5.Format.Borders.Color = Colors.Black;
            legenda5.Format.LeftIndent = leftIndent;
            legenda5.Format.Shading.Color = azzurro3;
     
        } // Legenda()


        void DefineStyles()
        {
            // Da rivedere se c'è qualcosa da modificare negli stili

            Style style = this.document.Styles["Normal"];
            style.Font.Name = "Arial";
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.Font.Size = 9;

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

    }
}
