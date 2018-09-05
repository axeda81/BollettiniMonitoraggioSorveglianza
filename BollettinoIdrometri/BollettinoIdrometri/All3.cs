using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace BollettiniMonitoraggio
{
    class All3
    {
        protected Document document;
        protected Table table;
        protected Section section;
        protected System.IO.StreamReader fileCSV;
        protected string nomeFileCSV; // Nome del CSV da aprire, passato come parametro
        protected string consultazione; // stringa che contiene ora e data della creazione del file - messa nel costruttore per impedire che possa scattare un minuto tra una pagina e l'altra 
        private LogWriter log; // collegamento al file di log dove scriverò informazioni sull'esecuzione
        protected string dirigente; // Nome del dirigente che firma il PDF, da prelevare da config.ini
        protected string filename;

        public All3(string n)
        {
            consultazione = "Estrazione dati delle ore " + DateTime.Now.ToString("HH:mm") + " del " + DateTime.Now.ToString("dd/MM/yyyy");
            log = LogWriter.Instance; // File di log in cui scrivo informazioni su esecuzione ed errori
            log.WriteToLog("ALL. 3 - Inizio scrittura bollettino, " + consultazione + ".", level.Info);
            filename = "Allegato3.pdf";
            dirigente = "";
            nomeFileCSV = n;

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
                log.WriteToLog("ALL. 3 - BollettinoPiogge: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception);
            }

            CreaPDFdaCSV();
            log.WriteToLog("ALL. 3 - Fine scrittura bollettino, " + consultazione + ".", level.Info);
        }

        private void CreaPDFdaCSV()
        {
            try
            {
                using (fileCSV = new StreamReader(nomeFileCSV)) // Apro il file CSV
                {
                    // Creazione del PDF
                    document = new Document();
                    document.Info.Title = "Bollettino altezze idrometriche";
                    document.Info.Author = "R.A.S.";

                    document.UseCmykColor = true;
                    const bool unicode = true;
                    const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

                    PdfDocument doc = new PdfDocument();

                    DefineStyles();
                    // Impaginazione delle informazioni nel pdf
                    CreatePage();

                    PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode, embedding); // opzioni di visualizzazione
                    pdfRenderer.Document = document;
                    pdfRenderer.RenderDocument();
                    //string filename = DateTime.Now.ToString("ALL. 3 - ddMMyyyy-HHmm") + ".pdf";

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

                    // Apre il pdf 
                    // System.Diagnostics.Process.Start(filename);
                } // using
            } //try

            catch (System.Exception err)
            {
                // Intercetta tutti i tipi di eccezione ma quelle che si dovrebbero verificare più di frequente sono:
                // System.IO.FileNotFoundException, System.IO.DirectoryNotFoundException, System.IO.IOException...        
                err.Source = "CreaPDFdaCSV()";
                // Salvo nel log il messaggio di errore con un pò di informazioni sulla funzione che ha lanciato l'eccezione e sul tipo di eccezione
                log.WriteToLog("ALL. 3 - GraficoIdrometri: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception);
            }
        }


        void CreatePage()
        {
            // Each MigraDoc document needs at least one section.
            section = this.document.AddSection();
            // Il pdf verrà stampato in orizzontale
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
            section.PageSetup.TopMargin = "10mm";
            section.PageSetup.LeftMargin = "05mm";
            section.PageSetup.RightMargin = "05mm";
            section.PageSetup.BottomMargin = "15mm";
            section.PageSetup.DifferentFirstPageHeaderFooter = false;

            section.PageSetup.FooterDistance = Unit.FromCentimeter(0); // distanza del footer dal fondo della pagina

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
            firma.Format.Alignment = ParagraphAlignment.Center;
            firma.Format.Font.Size = 7;
            firma.Format.Font.Bold = true;
            section.Footers.Primary.Add(firma);
            section.Footers.EvenPage.Add(firma.Clone());

            // ******************* Da qui in poi - creazione tabella *******************

            CreateTableHeader(); // Header (viene replicato in ogni pagina)

            CreateTableBody(); // Dati bollettino

            Legenda(); // Legenda sulla colorazione delle celle in caso di superamento soglie

            // Inserisco il disclaimer sui dati
            Paragraph disclaimer = new Paragraph();
            disclaimer.AddText("\n\n\n\n\"Composizione e rappresentazione dei dati eseguita con modalità automatiche su dati della rete di stazioni meteorologiche fiduciarie della Regione Sardegna gestita dall\'Agenzia per la Protezione dell'Ambiente della Sardegna, ARPAS, acquisiti in tempo reale e sottoposti ad un processo automatico di validazione di primo livello\"");
            disclaimer.Format.Font.Size = 6;
            disclaimer.Format.Font.Bold = false;
            disclaimer.Format.Alignment = ParagraphAlignment.Left;
            
            section.Add(disclaimer);

        } //CreatePage()        


        void CreateTableHeader()
        {
            // Crea l'intestazione della tabella che andrà ripetuta in ogni pagina
            table = section.AddTable();
            table.Borders.Width = 0.2;
            table.Rows.LeftIndent = 0;
            table.Format.Alignment = ParagraphAlignment.Center;
            const int numColonne = 11;
            table.TopPadding = 3;
            table.BottomPadding = 3;
            table.Format.Font.Size = 7;

            // Definisco le colonne
            Column column = table.AddColumn("3.2cm");
            column = table.AddColumn("2.8cm");
            column = table.AddColumn("3cm");
            column = table.AddColumn("3cm");
            column = table.AddColumn("4cm");
            column = table.AddColumn("1.5cm");
            column = table.AddColumn("1.5cm");
            column.Shading.Color = Colors.LightYellow;
            column = table.AddColumn("1.5cm");
            column.Shading.Color = Colors.LightYellow;
            column = table.AddColumn("1.5cm");
            column.Shading.Color = Colors.LightYellow;
            column = table.AddColumn("2cm");
            column = table.AddColumn("4.5cm");

            // Titolo
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Cells[0].MergeRight = 9;
            DateTime dateValue = DateTime.Now;
            String titolo = "\n\nALTEZZE IDROMETRICHE REGISTRATE DALLE STAZIONI DELLA RETE FIDUCIARIA";
            row.Cells[0].AddParagraph(titolo);
            row.Cells[0].AddParagraph(consultazione).Format.Font.Italic = true;
            row.Format.Font.Size = 9;
            row.Format.Font.Bold = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Borders.Visible = false;
            row.Cells[0].AddParagraph();

            MigraDoc.DocumentObjectModel.Shapes.Image image = new MigraDoc.DocumentObjectModel.Shapes.Image("./logo RAS con didascalia.png");
            image.LockAspectRatio = true;
            image.Height = Unit.FromCentimeter(2);
            image.Width = Unit.FromCentimeter(4.5);
            row.Cells[10].Add(image);

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Borders.Visible = false;
            row.Cells[0].MergeRight = 10;
            row.Height = 15; // questa riga serve solo a lasciare spazio prima dei dati

            // Intestazione tabella
            row = table.AddRow();
            row.Borders.Visible = true;
            row.HeadingFormat = true;
            row.Format.Font.Bold = true;
            row.Format.Font.Size = 7;

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

            Paragraph p = row.Cells[10].AddParagraph();
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

            for (int i = 0; i < numColonne; i++)
            {
                row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
            }

            row.Cells[10].Format.Alignment = ParagraphAlignment.Left;

        } // CreateTableHeader()


        void CreateTableBody()
        {
            // Fare prove con csv con formato non compatibile per vedere gestione eccezioni (teoricamente tutte gestite dal chiamante)

            string riga = fileCSV.ReadLine();
            string[] elencoCampi = { "" };

//            const int numRighePerPagina = 20; // num massimo di righe per pagina, poi si va a quella successiva
//            int numRigheScritte = 0; // contatore delle righe inserite 

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
                // Il file CSV era vuoto, lancio un'eccezione e lo segnalo nel log
                throw new System.IO.InvalidDataException("CreatePage(): File " + nomeFileCSV + " vuoto."); // Eccezione gestita dal chiamante
            }

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

                row.Format.Font.Bold = false;
                row.Cells[0].Shading.Color = Colors.LawnGreen;
                row.Cells[0].AddParagraph(elencoCampi[0]);
                row.Cells[0].MergeRight = 10;
                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                row.Cells[0].VerticalAlignment = VerticalAlignment.Center;

                // Poi ci sono le righe con dati di ogni stazione
                for (int s = 0; s < numStazioniPerBacino[b]; s++)
                {
                    riga = fileCSV.ReadLine();
                    elencoCampi = CSVRowToStringArray(riga, ';', '\n');
                    row = table.AddRow();

                    for (int i = 0; i < (elencoCampi.Length - 1); i++)
                    {
                        // Riempio tutta la riga con i valori del CSV, tranne l'ultima cella perchè 
                        // vanno inseriti dei simboli a seconda dell'entità della variazione
                        if (i <= 9) row.Cells[i].AddParagraph(elencoCampi[i]);

                        if (i == 0)
                        {
                            // Nel primo campo, oltre al nome della stazione, devo aggiungere l'informazione sull'ultimo dato disponibile,
                            // che trovo come ultimo campo nel CSV e quindi nell'ultimo elemento di elencoCampi
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
                            row.Cells[i].Add(p);
                        }

                        row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
                        row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                        if (i == 0) row.Cells[0].Format.Alignment = ParagraphAlignment.Left;

                        if (i == 9)
                        {
                            double s1, s2, s3, variazione;
                            double altezzaRegistrata = -1; // TODO trovare un altro modo di inizializzarlo perchè potrebbe pure essere negativo
                            variazione = -1; // Se rimane a -1 vuol dire che non c'erano dati nella riga, quindi non va scritto nessun simbolo nella cella corrispondente

                            if (elencoCampi[6] != "") s1 = Convert.ToDouble(elencoCampi[6]);
                            else s1 = 0;
                            if (elencoCampi[7] != "") s2 = Convert.ToDouble(elencoCampi[7]);
                            else s2 = 0;
                            if (elencoCampi[8] != "") s3 = Convert.ToDouble(elencoCampi[8]);
                            else s3 = 0;
                            if (elencoCampi[9] != "") altezzaRegistrata = Convert.ToDouble(elencoCampi[9]);

                            if (elencoCampi[10] != "") variazione = Convert.ToDouble(elencoCampi[10]);

                            if ((altezzaRegistrata <= s1) /*&& (altezzaRegistrata >= 0)*/)
                                row.Cells[9].Shading.Color = Colors.LimeGreen;
                            else if ((altezzaRegistrata > s1) && (altezzaRegistrata < s2) && (altezzaRegistrata >= 0))
                                row.Cells[9].Shading.Color = Colors.Yellow;
                            else if ((altezzaRegistrata >= s2) && (altezzaRegistrata <= s3) && (altezzaRegistrata >= 0))
                                row.Cells[9].Shading.Color = Colors.Orange;
                            else if ((altezzaRegistrata >= s3) && (altezzaRegistrata >= 0))
                                row.Cells[9].Shading.Color = Colors.Red;

                            if (variazione == 0)
                            {
                                row.Cells[10].AddParagraph("=");
                                row.Cells[10].Format.Font.Color = Colors.Blue;
                            }
                            else if (variazione == -1)
                            {
                                row.Cells[10].AddParagraph("");
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

                            row.Cells[10].Format.Font.Bold = true;
                            row.Cells[10].VerticalAlignment = VerticalAlignment.Center;
                            row.Cells[5].Format.Alignment = row.Cells[6].Format.Alignment = row.Cells[7].Format.Alignment = 
                                row.Cells[8].Format.Alignment = row.Cells[9].Format.Alignment = ParagraphAlignment.Center;

                        } // if (i == 9)
                    } // for (scorre gli elementi di una riga)
                } // for (scorre le stazioni di un bacino)
            } // for (scorre i bacini)

            if (numBacini == 0)
            {
                // Se il file CSV era vuoto aggiungo una riga alla tabella in cui lo scrivo
                Row row = table.AddRow();
                row.Cells[0].MergeRight = 10;
                row.Cells[0].Shading.Color = Colors.White;
                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                row.Cells[0].AddParagraph("Dati non presenti.");
            }

        } // CreateTableBody()

        void Legenda()
        {

            //Creo la legenda dove vengono indicati i criteri usati nella colorazione delle celle, da mettere a fine file 
            String nota = "Per la definizione delle soglie e per ulteriori informazioni verificare le monografie pubblicate nell\'apposita sezione del sito Internet della " +
                "Protezione Civile al link: http://www.sardegnaambiente.it/protezionecivile/nowcasting/monografie_idrometri.html\n";

            Paragraph spazio = section.AddParagraph();
            spazio.AddLineBreak(); // Lascio una riga di spazio fra i dati e la legenda

            int altezza = 12;
            Table legenda = document.LastSection.AddTable();
            legenda.Borders.Visible = false;
            legenda.AddColumn("19cm");
            legenda.AddColumn("1.5cm");
            legenda.AddColumn("8cm");
            legenda.Shading.Color = Colors.White;
            legenda.Rows.Height = 10;
            legenda.Rows.VerticalAlignment = VerticalAlignment.Center;
            legenda.Format.Font.Size = 7;

            Row row = legenda.AddRow();
            row.Height = altezza;
            row.Format.Font.Size = 6;
            row.Cells[0].MergeDown = 4;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].AddParagraph(nota);

            row.Cells[1].MergeRight = 1;
            row.Cells[1].AddParagraph("Legenda dei colori");
            row.Cells[1].Format.Font.Bold = false;
            row.Cells[1].Format.Font.Color = Colors.White;
            row.Cells[1].Shading.Color = Colors.Gray;
            row.Cells[1].Format.Alignment = ParagraphAlignment.Center;

            row = legenda.AddRow();
            row.Height = altezza;
            row.Format.Font.Size = 6;
            row.Cells[1].Shading.Color = Colors.LimeGreen;
            row.Cells[1].AddParagraph("h ≤ S1");
            row.Cells[2].AddParagraph("Livello idrometrico inferiore alla PRIMA soglia");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].Borders.Right.Visible = true;

            row = legenda.AddRow();
            row.Height = altezza;
            row.Format.Font.Size = 6;
            row.Cells[1].Shading.Color = Colors.Yellow;
            row.Cells[1].AddParagraph("S1 < h < S2");
            row.Cells[2].AddParagraph("Livello idrometrico compreso tra la PRIMA e la SECONDA soglia");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].Borders.Right.Visible = true;

            row = legenda.AddRow();
            row.Height = altezza;
            row.Format.Font.Size = 6;
            row.Cells[1].Shading.Color = Colors.Orange;
            row.Cells[1].AddParagraph("S2 ≤ h < S3");
            row.Cells[2].AddParagraph("Livello idrometrico compreso tra la SECONDA e la TERZA soglia");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].Borders.Right.Visible = true;

            row = legenda.AddRow();
            row.Height = altezza;
            row.Format.Font.Size = 6;
            row.Cells[1].Shading.Color = Colors.Red;
            row.Cells[1].AddParagraph("h ≥ S3");
            row.Cells[2].AddParagraph("Livello idrometrico superiore alla TERZA soglia");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;

            row.Cells[2].Borders.Right.Visible = true;
            row.Cells[2].Borders.Bottom.Visible = true;

        } // Legenda()


        void DefineStyles()
        {
            // Da rivedere se c'è qualcosa da modificare negli stili

            Style style = this.document.Styles["Normal"];
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

    }
}
