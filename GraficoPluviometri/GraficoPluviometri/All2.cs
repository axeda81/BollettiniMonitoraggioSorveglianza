using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using ZedGraph;


namespace BollettiniMonitoraggio
{
    class All2
    {
        protected Document document;
        protected Table table;
        protected Section section;
        protected StreamReader fileCSV;
        protected int numStazioni = 0;
        protected DateTime inizio;
        protected string nomeFileCSV; // Nome del CSV da aprire, passato come parametro
        protected string consultazione; // stringa che contiene ora e data della creazione del file - messa nel costruttore per impedire che possa scattare un minuto tra una pagina e l'altra 
        private LogWriter log; // collegamento al file di log dove scriverò informazioni sull'esecuzione
        private const string finefileBmp = "_2.bmp";
        protected string dirigente; // Nome del dirigente che firma il PDF, da prelevare da config.ini
        protected string filename;

        public All2(string n, string d, string t)
        {
            consultazione = "Estrazione dati delle ore " + DateTime.Now.ToString("HH:mm") + " del " + DateTime.Now.ToString("dd/MM/yyyy");
            log = LogWriter.Instance; // File di log in cui scrivo informazioni su esecuzione ed errori
            log.WriteToLog("ALL. 2 - Inizio scrittura bollettino, " + consultazione + ".", level.Info);
            string s = d + " " + t;
            inizio = DateTime.Parse(s);
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
                log.WriteToLog("ALL. 2 - BollettinoPiogge: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception);
            }

            filename = "Allegato2.pdf";
            nomeFileCSV = n;
            CreaPDFdaCSV();
            log.WriteToLog("ALL. 2 - Fine scrittura bollettino, " + consultazione + ".", level.Info);

        }

        private void CreaPDFdaCSV()
        {
            try
            {
                using (fileCSV = new StreamReader(nomeFileCSV))
                {
                    // Creazione del PDF
                    document = new Document();
                    document.Info.Title = "Pioggia registrata nelle ultime 24 ore";
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

                    // A questo punto le Bitmap devono poter essere cancellate
                    eliminaBitmap();

                    //string filename = "ALL. 2 - " + DateTime.Now.ToString("ddMMyyyy-HHmm") + ".pdf";

                    // Alla fine salva il pdf appena creato: se esiste già con lo stesso nome lo elimina e poi salva quello nuovo
                    if (System.IO.File.Exists(filename))
                    {
                        FileInfo bInfoOld = new FileInfo(filename);
                        bInfoOld.IsReadOnly = false;
                        System.IO.File.Delete(filename);
                    }

                    pdfRenderer.PdfDocument.Save(filename);

                    // Imposto il file come accessibile in sola lettura
                    FileInfo bInfo = new FileInfo(filename);
                    bInfo.IsReadOnly = true;

                    // System.Diagnostics.Process.Start(filename);
                } //using
            } //try

            catch (System.Exception err)
            {
                if (numStazioni > 0) eliminaBitmap(); // Elimino eventuali bmp che erano state già create
                // Intercetta tutti i tipi di eccezione ma quelle che si dovrebbero verificare più di frequente sono:
                // System.IO.FileNotFoundException, System.IO.DirectoryNotFoundException, System.IO.IOException...        
                err.Source = "CreaPDFdaCSV()";
                // Salvo nel log il messaggio di errore con un pò di informazioni sulla funzione che ha lanciato l'eccezione e sul tipo di eccezione
                log.WriteToLog("ALL. 2 - GraficoPluviometri: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception);
            }
        }

        private void CreatePage()
        {
            // Costruzione pdf da fare con FOR (num stazioni) 
            // Per ogni stazione una pagina, composta da Header (logo, titolo...) e grafico con i dati
            // Dati per una stazione = 1 riga del csv
            string riga = fileCSV.ReadLine();
            string[] elencoCampi = { "" };

            while ((riga != null) && (riga != ""))
            {
                // Conto quante righe ci sono nel file (e quindi quante stazioni)
                numStazioni++;
                riga = fileCSV.ReadLine();
            }

            if (numStazioni == 0)
            {
                // Il file CSV era vuoto, lancio un'eccezione e lo segnalo nel log
                throw new System.IO.InvalidDataException("CreatePage(): File " + nomeFileCSV + " vuoto."); // Eccezione gestita dal chiamante
            }

            // Each MigraDoc document needs at least one section.
            section = document.AddSection();
            // Il pdf verrà stampato in orizzontale
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
            section.PageSetup.TopMargin = "05mm";
            section.PageSetup.LeftMargin = "05mm";
            section.PageSetup.RightMargin = "05mm";
            section.PageSetup.BottomMargin = "05mm";

            section.PageSetup.FooterDistance = Unit.FromCentimeter(0.5);

            // Aggiungo i numeri di pagina
            section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            Paragraph paragraph = new Paragraph();
            paragraph.Format.Font.Size = 7;
            paragraph.Format.Alignment = ParagraphAlignment.Right;
            paragraph.AddText("Pagina ");
            paragraph.AddPageField();
            paragraph.AddText(" di ");
            paragraph.AddNumPagesField();
            paragraph.AddText("\n");
            section.Footers.Primary.Add(paragraph);
            section.Footers.EvenPage.Add(paragraph.Clone());
            // Aggiungo il disclaimer sui dati nel footer
            Paragraph disclaimer = new Paragraph();
            disclaimer.Format.Font.Bold = false;
            disclaimer.Format.Font.Size = 6;
            disclaimer.Format.Alignment = ParagraphAlignment.Left;
            FormattedText fd = disclaimer.AddFormattedText("\"Composizione e rappresentazione dei dati eseguita con modalità automatiche su dati della rete di stazioni meteorologiche fiduciarie della Regione Sardegna gestita dall\'Agenzia per la Protezione dell'Ambiente della Sardegna, ARPAS,\nacquisiti in tempo reale e sottoposti ad un processo automatico di validazione di primo livello\"");
            section.Footers.Primary.Add(disclaimer);
            section.Footers.EvenPage.Add(disclaimer.Clone());

            fileCSV.BaseStream.Seek(0, SeekOrigin.Begin); // Torno a inizio file csv

            for (int i = 0; i < numStazioni; i++)
            {
                riga = fileCSV.ReadLine();

                // Elimino il primo carattere del CSV, se è il char di controllo della codifica 
                if ((i == 0) && (riga.ToCharArray()[0] == 65279))
                    riga = riga.Remove(0, 1);

                elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                CreateHeader(elencoCampi[0], elencoCampi[1], elencoCampi[elencoCampi.Length - 1]);
                DisegnaGrafico(elencoCampi, i);

            } // for (numStazioni)

            //if (numStazioniConPioggia == 0)
            //{
            //    // il file csv non era vuoto, ma i valori erano tutti a 0 - lo scrivo nel log
            //    log.WriteToLog("ALL. 2 - Non sono presenti piogge su nessuna stazione", level.Info);
            //    // Creo comunque un file pdf, ma senza grafici - NO, vanno comunque disegnati tutti i grafici
            //    //CreateEmptyHeader();
            //    //CreateEmptyBody();
            //}

        } //CreatePage()        


        private void CreateHeader(string nomeStazione, string nomeArea, string ultimoDatoDisponibile)
        {
            // Le informazioni di intestazione le metto in una tabella 
            table = section.AddTable();
            Column column = table.AddColumn("17cm");
            column = table.AddColumn("5cm");
            column = table.AddColumn("5cm");

            Row row = table.AddRow();
            row.Borders.Visible = false;

            // Titolo
            DateTime dateValue = DateTime.Now;  
            string stazione = "Stazione pluviometrica " + nomeStazione + "\nArea " + nomeArea;
            string titolo = "Pioggia registrata nelle ultime 24 ore";

            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].AddParagraph(stazione).Format.Font.Bold = true;
            row.Cells[0].AddParagraph(titolo).Format.Font.Bold = true;
            row.Cells[0].AddParagraph(consultazione).Format.Font.Italic = true;

            // Devo aggiungere l'informazione sull'ultimo dato disponibile, che trovo come ultimo campo nel CSV e quindi nell'ultimo elemento di elencoCampi
            Paragraph p = new Paragraph();
            p.AddFormattedText("Ultimo dato disponibile: ");

            // aggiunta di data e ora ultimo dato disponibile, prendendo da csv stringa formattata come ddMMyyyyHHmm
            ultimoDatoDisponibile = ultimoDatoDisponibile.Insert(2, "/");
            ultimoDatoDisponibile = ultimoDatoDisponibile.Insert(5, "/");
            ultimoDatoDisponibile = ultimoDatoDisponibile.Insert(10, " alle ");
            ultimoDatoDisponibile = ultimoDatoDisponibile.Insert(18, ":");
            p.AddFormattedText(ultimoDatoDisponibile);

            p.Format.Font.Italic = true;
            p.Format.Font.Size = 6;
            row.Cells[0].Add(p);

            MigraDoc.DocumentObjectModel.Shapes.Image image = new MigraDoc.DocumentObjectModel.Shapes.Image("./logo RAS con didascalia.png");
            image.LockAspectRatio = true;
            image.Height = Unit.FromCentimeter(2);
            image.Width = Unit.FromCentimeter(4.5);
            row.Cells[1].Add(image);

            row.Cells[2].AddParagraph("\nARPAS\nF.to il Dirigente Responsabile\n" + dirigente);
            row.Cells[2].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[2].Format.Font.Size = 7;
            row.Cells[2].Format.Font.Bold = true;

            row = table.AddRow();
            row.Borders.Visible = false;
            row.Height = 5; // questa riga serve solo a lasciare spazio prima del grafico
        }

        private void DefineStyles()
        {
            // Da rivedere se c'è qualcosa da modificare negli stili
            Style style = this.document.Styles["Normal"];
            style.Font.Name = "Arial";
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.Font.Size = 9;

        } //DefineStyles()


        private void DisegnaGrafico(string[] elencoCampi, int index)
        {
            // Creo un grafico con la libreria ZedGraph (devo usarla perchè mi permette di creare grafici con due assi Y in scale diverse) 
            // Poi lo salvo come BMP e carico la BMP nel PDF 

            const int fontSize = 9;
            ZedGraphControl graphControl = new ZedGraphControl();
            ZedGraph.GraphPane chart = graphControl.GraphPane;

            // Impostazioni generali del grafico
            chart.Title.Text = "";
            chart.Border.IsVisible = false;
            chart.XAxis.Title.Text = "Tempo (ore)";
            chart.YAxis.Title.Text = "h (mm in 15')";
            chart.Y2Axis.Title.Text = "h cumulata (mm)";
            chart.XAxis.Title.FontSpec.IsBold = chart.YAxis.Title.FontSpec.IsBold = chart.Y2Axis.Title.FontSpec.IsBold = true;
            chart.XAxis.Title.FontSpec.Size = chart.YAxis.Title.FontSpec.Size = chart.Y2Axis.Title.FontSpec.Size = fontSize;
            chart.XAxis.Scale.FontSpec.Size = chart.X2Axis.Scale.FontSpec.Size = chart.YAxis.Scale.FontSpec.Size = chart.Y2Axis.Scale.FontSpec.Size = fontSize;

            chart.Legend.FontSpec.Size = fontSize;
            chart.Legend.Position = LegendPos.InsideTopLeft;
            chart.Legend.Border.Color = System.Drawing.Color.SteelBlue;
            chart.Legend.Border.Width = 3;
            chart.Legend.IsVisible = true;

            // Impostazioni asse X
            chart.XAxis.Type = AxisType.Text;
            chart.XAxis.Scale.MajorStep = 4;
            chart.XAxis.Scale.MinorStep = 1;
            chart.XAxis.MajorGrid.IsVisible = true;
            chart.XAxis.MinorGrid.IsVisible = true;
            chart.XAxis.MinorGrid.PenWidth = 1;
            chart.XAxis.MajorGrid.Color = chart.XAxis.MinorGrid.Color = System.Drawing.Color.Black;
            chart.XAxis.MajorGrid.DashOff = 0;
            chart.XAxis.MinorTic.Size = 1;
            chart.XAxis.Scale.FontSpec.Angle = 270;
            chart.XAxis.Scale.Min = 1;

            // Creo un array di double in cui mettere tutti i valori presi dal csv
            // E calcolo tutti i valori di pioggia cumulata che poi andranno rappresentati tramite linea
            // Questi array devono avere due campi in meno rispetto al csv perchè i primi due campi hanno nome stazione e area
            double[] pioggiaRegistrata = new double[elencoCampi.Length - 3];
            double[] pioggiaCumulata = new double[elencoCampi.Length - 3];

            double[] valoriAsseXprova = new double[elencoCampi.Length - 3];
            double[] valoriAsseYprova = new double[elencoCampi.Length - 3];

            string[] valoriAsseX = new string[elencoCampi.Length - 3];

            DateTime ieri = inizio.AddHours(-24);
            bool cambiaGiorno = false;

            valoriAsseX[0] = ieri.ToString("HH:mm\ndd MMM yyyy");

            for (int i = 0; i < pioggiaRegistrata.Length; i++)
            {
                // Prima di tutto, valutare se il dato non è disponibile 
                if ((elencoCampi[i + 2] == "") || (elencoCampi[i + 2].Contains("N/D")))
                {
                    // Dato non disponibile, quindi non devo esaminarlo e convertirlo in double
                    pioggiaRegistrata[i] = 0;
                    if (i > 0)
                        pioggiaCumulata[i] = pioggiaCumulata[i - 1]; // Asse Y dx, il grafico rimane costante
                    else pioggiaCumulata[i] = 0;
                    // Salvo le coordinate del valore N/D per la curva nd che poi disegnerò
                    valoriAsseXprova[i] = i;
                    valoriAsseYprova[i] = pioggiaCumulata[i]; 
                }
                else
                {
                    // Dato disponibile, lo converto in double e lo aggiungo anche alla pioggia cumulata
                    pioggiaRegistrata[i] = Double.Parse(elencoCampi[i + 2]); // Asse Y sx 
                    if (i == 0) pioggiaCumulata[i] = pioggiaRegistrata[i]; // Asse Y dx
                    else pioggiaCumulata[i] = pioggiaCumulata[i - 1] + pioggiaRegistrata[i]; // Asse Y dx
                }

                // *** ETICHETTE ASSE X *** un'etichetta ogni quarto d'ora (ne verrà visualizzata una per ogni ora)            
                int tmp = ieri.Day;
                DateTime tmpDate = ieri;
                ieri = ieri.AddMinutes(15);

                if (tmp != ieri.Day)
                    cambiaGiorno = true;
                if ((cambiaGiorno) && (i % 4 == 0))
                {
                    valoriAsseX[i] = tmpDate.ToString("HH:mm\ndd MMM yyyy");
                    cambiaGiorno = false;
                }
                else
                    valoriAsseX[i] = tmpDate.ToString("HH:mm");
              }

            // Metto le etichette con data e orari sull'asse x in basso
            chart.XAxis.Scale.TextLabels = valoriAsseX;

            // Elenco dei punti N/D
            LineItem nd = chart.AddCurve("Valori mancanti", valoriAsseXprova, valoriAsseYprova, System.Drawing.Color.DarkGray);
            nd.IsY2Axis = true;
            nd.IsX2Axis = true;
            nd.Line.Width = 0;
            nd.Line.IsVisible = false;
            nd.Symbol.IsVisible = true;
            nd.Symbol.Size = 2;
            nd.Symbol.Fill.Type = FillType.Solid;
            nd.Symbol.Type = SymbolType.Square;
            chart.AxisChange(); 

            // Curva della pioggia cumulata 
            LineItem curva = chart.AddCurve("Pioggia cumulata", null, pioggiaCumulata, System.Drawing.Color.Red);
            curva.Symbol.IsVisible = true;
            curva.Symbol.Size = 1;
            curva.Symbol.Fill.Type = FillType.Solid;
            curva.Line.Width = 7; 
            // Questa seconda curva viene associata all'asse Y2
            curva.IsY2Axis = true;
            curva.IsX2Axis = true;

            // Istogramma della pioggia misurata sull'asse Y
            BarItem istogramma = chart.AddBar("Pioggia registrata\n", null, pioggiaRegistrata, System.Drawing.Color.SteelBlue);
            istogramma.IsVisible = true;
            istogramma.Bar.Border.Color = System.Drawing.Color.SteelBlue;
            istogramma.Bar.Fill.Type = FillType.Solid;
            istogramma.IsX2Axis = true;

            // Impostazioni asse Y (grafico a barre)
            chart.YAxis.MajorTic.IsOpposite = false;
            chart.YAxis.MinorTic.IsOpposite = false;
            chart.YAxis.MajorGrid.Color = System.Drawing.Color.Black;
            chart.YAxis.MajorGrid.IsVisible = true;
            chart.YAxis.MajorGrid.DashOff = 0;
            chart.YAxis.MajorGrid.PenWidth = 1;
            chart.YAxis.MinorTic.Size = 1;
            chart.YAxis.MajorTic.Size = 3;
            chart.YAxis.MajorGrid.Color = chart.YAxis.MinorGrid.Color = System.Drawing.Color.Black;
            chart.XAxis.Title.FontSpec.IsAntiAlias = chart.XAxis.Scale.FontSpec.IsAntiAlias = chart.YAxis.Title.FontSpec.IsAntiAlias = chart.Legend.FontSpec.IsAntiAlias = true;
            chart.YAxis.Scale.Align = AlignP.Inside;

            // Impostazioni asse Y2 (curva pioggia cumulata)
            chart.Y2Axis.IsVisible = true;
            chart.Y2Axis.MajorTic.IsOpposite = false;
            chart.Y2Axis.MinorTic.IsOpposite = false;
            chart.Y2Axis.MinorTic.Size = 1;
            chart.Y2Axis.MajorTic.Size = 3;
            chart.Y2Axis.MajorGrid.Color = chart.Y2Axis.MinorGrid.Color = System.Drawing.Color.Black;
            chart.Y2Axis.Scale.Align = AlignP.Inside;

            // Asse X2 lineare 
            chart.X2Axis.Type = AxisType.Linear;
            chart.X2Axis.IsVisible = false;
            chart.X2Axis.MajorGrid.IsVisible = false;
            chart.X2Axis.MinorGrid.IsVisible = false;
            chart.X2Axis.Scale.Max = valoriAsseXprova.Length;
            chart.X2Axis.Scale.Min = 1;

            // graphControl.Invalidate();
            chart.AxisChange(); // Fa in modo che il grafico venga ridisegnato 

            String percorsoImmagine = "./" + index.ToString() + finefileBmp;
            // Riga vuota per lasciare spazio
            Row row = this.table.AddRow();
            row.Borders.Visible = false;
            row.Height = 25;

            // Salvo il grafico come immagine e poi lo carico nel pdf
            graphControl.MasterPane.ReSize(graphControl.CreateGraphics(), new RectangleF(0, 0, 2000, 1200));
            Bitmap b = graphControl.GraphPane.GetImage();

            // Se la bitmap esiste già (se ad esempio l'esecuzione precedente non è andata a buon fine e non
            // sono state cancellate correttamente), prima la cancello e poi la ricreo
            if (File.Exists(percorsoImmagine))
            {
                FileInfo fi = new FileInfo(percorsoImmagine);
                fi.IsReadOnly = false;
                File.Delete(percorsoImmagine);
            } 
                            
            b.Save(percorsoImmagine);
            // Le bitmap devono essere ReadOnly, almeno finchè non si finisce di confezionare il PDF
            FileInfo fInfo = new FileInfo(percorsoImmagine);
            fInfo.IsReadOnly = true;

            MigraDoc.DocumentObjectModel.Shapes.Image image = new MigraDoc.DocumentObjectModel.Shapes.Image(percorsoImmagine);
            image.Width = "750pt";
            image.Height = "425pt";
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Center;

            section.Add(image);
            if (index != (numStazioni - 1)) section.AddPageBreak();
            b.Dispose();
        }

        //private void CreateEmptyHeader()
        //{
        //    // Da chiamare quando non devo disegnare nessun grafico perchè non ci sono piogge da nessuna parte 

        //    // Le informazioni di intestazione le metto in una tabella 
        //    table = section.AddTable();
        //    Column column = table.AddColumn("17cm");
        //    column = table.AddColumn("10cm");

        //    Row row = table.AddRow();
        //    row.Borders.Visible = false;

        //    MigraDoc.DocumentObjectModel.Shapes.Image image = new MigraDoc.DocumentObjectModel.Shapes.Image("./logo RAS con didascalia.png");
        //    image.LockAspectRatio = true;
        //    image.Height = Unit.FromCentimeter(2);
        //    image.Width = Unit.FromCentimeter(4.5);
        //    row.Cells[0].Add(image);

        //    row.Cells[1].AddParagraph("\nARPAS\nF.to il Dirigente Responsabile\n" + dirigente);
        //    row.Cells[1].Format.Alignment = ParagraphAlignment.Right;
        //    row.Cells[1].Format.Font.Size = 7;
        //    row.Cells[1].Format.Font.Bold = true;

        //    row = table.AddRow();
        //    row.Borders.Visible = false;
        //    row.Height = 30; // questa riga serve solo a lasciare spazio prima del grafico
        //}

        //private void CreateEmptyBody()
        //{
        //    Row row = table.AddRow();
        //    row.Borders.Visible = false;
            
        //    // Titolo
        //    DateTime dateValue = DateTime.Now;
        //    string titolo = "Nessuna pioggia registrata nelle ultime 24 ore";
        //    row.Cells[0].MergeRight = 1;
        //    row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
        //    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
        //    row.Cells[0].AddParagraph(titolo).Format.Font.Bold = true;
        //    row.Cells[0].AddParagraph(consultazione).Format.Font.Italic = true;

        //    row = table.AddRow();
        //    row.Cells[0].MergeRight = 1;
        //    row.Borders.Visible = false;
        //    row.Height = 30;

        //    // Inserisco il disclaimer sui dati
        //    string disclaimer = "\n\n\"Composizione e rappresentazione dei dati eseguita con modalità automatiche su dati della rete di stazioni meteorologiche fiduciarie della Regione Sardegna gestita dall\'Agenzia per la Protezione dell'Ambiente della Sardegna, ARPAS, acquisiti in tempo reale e sottoposti ad un processo automatico di validazione di primo livello\"";
        //    row = table.AddRow();
        //    row.Cells[0].MergeRight = 1;
        //    row.Borders.Visible = false;
        //    row.Cells[0].AddParagraph(disclaimer);
        //    row.Format.Font.Size = 6;
        //    row.Format.Alignment = ParagraphAlignment.Left;
        //}

        private void eliminaBitmap()
        {
            // Elimina le Bitmap che sono state create durante l'elaborazione 
            for (int i = 0; i < numStazioni; i++)
            {
                string filebmp = i.ToString() + finefileBmp;
                if (File.Exists(filebmp))
                {
                    FileInfo fInfo = new FileInfo(filebmp);
                    fInfo.IsReadOnly = false;
                    File.Delete(filebmp);
                }
            }
        }

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
