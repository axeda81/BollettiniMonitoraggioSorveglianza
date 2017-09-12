using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ZedGraph;
using System.Drawing;
using System.Windows.Forms;

namespace BollettiniMonitoraggio
{
    class All4
    {
        // Indici delle soglie e dell'altezza misurata all'interno del vettore elencoCampi in cui ogni volta viene salvata una riga del CSV
        const int indiceS1 = 3;
        const int indiceS2 = 4;
        const int indiceS3 = 5;
        const int indiceH = 6;

        protected Document document;
        protected Table table;
        protected Section section;
        protected StreamReader fileCSV;
        protected int numStazioni = 0;
        protected string nomeFileCSV; // Nome del CSV da aprire, passato come parametro
        protected string consultazione; // stringa che contiene ora e data della creazione del file - messa nel costruttore per impedire che possa scattare un minuto tra una pagina e l'altra 
        private const string finefileBmp = "_4.bmp";
        protected string dirigente; // Nome del dirigente che firma il PDF, da prelevare da config.ini
        protected string filename;

        private LogWriter log; // collegamento al file di log dove scriverò informazioni sull'esecuzione

        public All4 (string n)
        {
            consultazione = "Estrazione dati delle ore " + DateTime.Now.ToString("HH:mm") + " del " + DateTime.Now.ToString("dd/MM/yyyy");
            log = LogWriter.Instance; // File di log in cui scrivo informazioni su esecuzione ed errori
            log.WriteToLog("ALL. 4 - Inizio scrittura bollettino, " + consultazione + ".", level.Info);
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
                log.WriteToLog("ALL. 4 - BollettinoPiogge: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception);
            }

            filename = "Allegato4.pdf";
            nomeFileCSV = n;
            CreaPDFdaCSV();
            log.WriteToLog("ALL. 4 - Fine scrittura bollettino, " + consultazione + ".", level.Info);
        }

        private void CreaPDFdaCSV()
        {
            try
            {
                using (fileCSV = new StreamReader(nomeFileCSV, true))
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

                    // Alla fine salva il pdf appena creato: se esiste già con lo stesso nome lo elimina e poi salva quello nuovo
                    if (System.IO.File.Exists(filename))
                    {
                        FileInfo bInfoOld = new FileInfo(filename);
                        bInfoOld.IsReadOnly = false;
                        System.IO.File.Delete(filename);
                    }

                    pdfRenderer.PdfDocument.Save(filename);
                    // Imposto il file finale come accessibile in sola lettura
                    FileInfo bInfo = new FileInfo(filename);
                    bInfo.IsReadOnly = true;
                    // System.Diagnostics.Process.Start(filename);

                } // using(fileCSV): alla fine dovrebbe chiudere lo stream e rilasciare il file
            } // try

            catch (System.Exception err)
            {
                if (numStazioni > 0) eliminaBitmap(); // Elimino eventuali bmp che erano state già create
                // Intercetta tutti i tipi di eccezione ma quelle che si dovrebbero verificare più di frequente sono:
                // System.IO.FileNotFoundException, System.IO.DirectoryNotFoundException, System.IO.IOException...        
                err.Source = "CreaPDFdaCSV()"; 
                // Salvo nel log il messaggio di errore con un pò di informazioni sulla funzione che ha lanciato l'eccezione e sul tipo di eccezione
                log.WriteToLog("ALL. 4 - GraficoIdrometri: eccezione di tipo " + err.GetType().ToString() + " (" + err.Message + ")", level.Exception); 
            }
        }

        private void CreatePage()
        {
            // Costruzione pdf da fare con FOR (num idrometri) 
            // Per ogni idrometro una pagina, composta da Header e grafico con i dati
            // Dati per un idrometro = 1 riga del csv

            string riga = fileCSV.ReadLine();
            string[] elencoCampi = { "" };

            while ((riga != null) && (riga != ""))
            {
                // Conto quante righe ci sono nel file (e quindi quanti idrometri)
                numStazioni++;
                riga = fileCSV.ReadLine();
            }

            if (numStazioni == 0)
            {
                // Il file CSV era vuoto, lancio un'eccezione e lo segnalo nel log
                throw new System.IO.InvalidDataException("CreatePage(): File " + nomeFileCSV +" vuoto."); // Eccezione gestita dal chiamante
            }

            // Each MigraDoc document needs at least one section.
            section = this.document.AddSection();
            // Il pdf verrà stampato in orizzontale
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
            section.PageSetup.TopMargin = "05mm";
            section.PageSetup.LeftMargin = "05mm";
            section.PageSetup.RightMargin = "05mm";
            section.PageSetup.BottomMargin = "05mm";

            section.PageSetup.FooterDistance = Unit.FromCentimeter(0.5);

            // Aggiungo i numeri di pagina
            section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            Paragraph numPag = new Paragraph();
            numPag.Format.Font.Size = 7;
            numPag.Format.Alignment = ParagraphAlignment.Right;
            numPag.AddText("Pagina ");
            numPag.AddPageField();
            numPag.AddText(" di ");
            numPag.AddNumPagesField();
            section.Footers.Primary.Add(numPag);
            section.Footers.EvenPage.Add(numPag.Clone());

            // In fondo al documento aggiungo la firma del dirigente
            Paragraph firma = new Paragraph();
            firma.AddFormattedText("ARPAS\nF.to il Dirigente Responsabile\n" + dirigente);
            firma.Format.Alignment = ParagraphAlignment.Center;
            firma.Format.Font.Size = 7;
            firma.Format.Font.Bold = true;
            section.Footers.Primary.Add(firma);
            section.Footers.EvenPage.Add(firma.Clone());

            fileCSV.BaseStream.Seek(0, SeekOrigin.Begin); // Torno a inizio file csv

            for (int i = 0; i < numStazioni; i++)
            {
                riga = fileCSV.ReadLine();

                // Elimino il primo carattere del CSV, se è il char di controllo della codifica 
                // (dovrebbe essere UTF-8 e invece è UTF-16 Big Endian)
                if ((i == 0) && (riga.ToCharArray()[0] == 65279))
                    riga = riga.Remove(0, 1);

                elencoCampi = CSVRowToStringArray(riga, ';', '\n');

                // Riordino le coordinate dei punti: sono ordinate per X ma non per Y: se ci sono due punti 
                // con stessa X c'è il rischio che siano nell'ordine sbagliato
                double[] puntiOrdinati = riordinaPunti(elencoCampi);

                // Creo la pagina relativa all'i-esimo idrometro
                CreateHeader(elencoCampi[0], elencoCampi[1], elencoCampi[elencoCampi.Length - 1]);
                DisegnaGrafico(elencoCampi, puntiOrdinati, i);
            }

        } //CreatePage()        

        private void CreateHeader(string nomeStazione, string ubicazione, string ultimoDatoDisponibile)
        {
            // Le informazioni di intestazione le metto in una tabella 
            table = section.AddTable();
            table.Borders.Visible = false;
            Column column = table.AddColumn("23cm");
            column = table.AddColumn("3cm");
            column = table.AddColumn("2cm");
            Row row = table.AddRow();

            // Titolo
            DateTime dateValue = DateTime.Now;
            string titolo = "Altezza idrometrica registrata ";
            string stazione = "Stazione " + nomeStazione;

            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].AddParagraph(titolo).Format.Font.Bold = true;
            row.Cells[0].AddParagraph(stazione).Format.Font.Bold = true;
            row.Cells[0].AddParagraph(ubicazione);
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
            row.Cells[1].MergeRight = 1;
            row.Cells[1].Add(image);

            row = table.AddRow();
            row.Height = 5; // serve solo a lasciare spazio prima del grafico
        }

        private void DisegnaGrafico(String[] elencoCampi, double[] punti, int index)
        {
            // Le informazioni sulle soglie e sull'altezza misurata servono per disegnare delle rette a quei livelli
            double s1 = 0, s2 = 0, s3 = 0, h = 0;
            s1 = Convert.ToDouble(elencoCampi[indiceS1]);
            s2 = Convert.ToDouble(elencoCampi[indiceS2]);
            s3 = Convert.ToDouble(elencoCampi[indiceS3]);
            h = Convert.ToDouble(elencoCampi[indiceH]);

            int valMaxX = (int)(punti[punti.Length - 2]); // Valore massimo di X che devo rappresentare

            Double[] X = new Double[punti.Length / 2];
            Double[] Y = new Double[punti.Length / 2];

            // Preparo un array di stringhe in cui mettere i valori delle soglie che andranno disegnati sul grafico sul lato destro
            string[] valoriAsseY2 = new string[4];
            Double[] valoriAsseY2numerici = new Double[4];
            valoriAsseY2numerici[0] = s1;
            valoriAsseY2numerici[1] = s2;
            valoriAsseY2numerici[2] = s3;
            valoriAsseY2numerici[3] = h;
            Array.Sort(valoriAsseY2numerici);
            for (int v = 0; v < 4; v++) valoriAsseY2[v] = valoriAsseY2numerici[v].ToString();

            // Metto nei vettori X e Y le coordinate dei punti della spezzata che devo disegnare
            int cont = 0;
            for (int i = 0; i < (punti.Length / 2); i++)
            {
                X[i] = punti[i + cont];
                Y[i] = punti[i + 1 + cont];
                cont++;
            }

            ZedGraphControl graphControl = new ZedGraphControl();
            ZedGraph.GraphPane chart = graphControl.GraphPane;

            // Impostazioni generali del grafico
            chart.Title.Text = "";
            chart.Border.IsVisible = false;
            chart.XAxis.Title.Text = "Distanze (m)";
            chart.XAxis.Title.FontSpec.Size = chart.XAxis.Scale.FontSpec.Size = chart.YAxis.Title.FontSpec.Size = 
                chart.YAxis.Scale.FontSpec.Size = chart.Y2Axis.Scale.FontSpec.Size = 9;
            chart.Legend.FontSpec.Size = 8;
            chart.XAxis.Title.FontSpec.IsBold = chart.XAxis.Scale.FontSpec.IsBold = chart.YAxis.Title.FontSpec.IsBold = 
                chart.Y2Axis.Scale.FontSpec.IsBold = chart.Legend.FontSpec.IsBold = false;
            chart.XAxis.Title.FontSpec.IsAntiAlias = chart.XAxis.Scale.FontSpec.IsAntiAlias = chart.YAxis.Title.FontSpec.IsAntiAlias = 
                chart.Y2Axis.Title.FontSpec.IsAntiAlias = chart.Legend.FontSpec.IsAntiAlias = true;
            chart.XAxis.Type = AxisType.Linear;

            chart.XAxis.Scale.MajorStep = 5;
            //chart.XAxis.Scale.MinorStep = 1;
            chart.XAxis.MajorGrid.IsVisible = true;
            chart.XAxis.MinorGrid.IsVisible = false;
            chart.XAxis.MinorGrid.PenWidth = 0.5f;
            chart.XAxis.MajorGrid.Color = System.Drawing.Color.LightGray;
            chart.XAxis.MajorGrid.DashOff = 0;
            chart.XAxis.MinorTic.Size = 1;
            chart.XAxis.Scale.Max = valMaxX + 1.5F;
            chart.XAxis.Scale.Min = -0.5F;

            chart.YAxis.Title.Text = "h (m)";
            chart.YAxis.Scale.MajorStep = 5;
            chart.YAxis.Scale.MinorStep = 1;
            chart.YAxis.MajorGrid.IsVisible = true;
            chart.YAxis.MinorGrid.IsVisible = false;
            chart.YAxis.MinorGrid.PenWidth = 0.5f;
            chart.YAxis.MajorGrid.Color = System.Drawing.Color.LightGray;
            chart.YAxis.MajorGrid.DashOff = 0;
            chart.YAxis.MinorTic.Size = 1;

            LineItem curva = chart.AddCurve("", X, Y, System.Drawing.Color.Black); // spezzata 
            curva.Symbol.Size = 0.5F;
            curva.Line.Width = 5;

            // Vettori che mi servono per definire i due punti necessari per tracciare le rette alle altezze S1, S2, S3, h
            double[] tmpX = new double[2];
            double[] tmpY = new double[2];
            string label = "";
            tmpX[0] = 0;
            tmpX[1] = valMaxX + 1;
            tmpY[0] = s1;
            tmpY[1] = s1;
            label = "S1 = " + s1.ToString();
            curva = chart.AddCurve(label, tmpX, tmpY, System.Drawing.Color.Yellow);
            curva.Symbol.Size = 0.5F;
            curva.Line.Width = 5;
            tmpY[0] = tmpY[1] = s2;
            label = "S2 = " + s2.ToString();
            curva = chart.AddCurve(label, tmpX, tmpY, System.Drawing.Color.Orange);
            curva.Symbol.Size = 0.5F;
            curva.Line.Width = 5;
            tmpY[0] = tmpY[1] = s3;
            label = "S3 = " + s3.ToString();
            curva = chart.AddCurve(label, tmpX, tmpY, System.Drawing.Color.Red);
            curva.Symbol.Size = 0.5F;
            curva.Line.Width = 5;
            tmpY[0] = tmpY[1] = h;
            label = "h = " + h.ToString();
            curva = chart.AddCurve(label, tmpX, tmpY, System.Drawing.Color.Blue);
            curva.Symbol.Size = 0.5F;
            curva.Line.Width = 5;

            Legend legend = chart.Legend;
            legend.IsVisible = true;
            legend.Border.IsVisible = true;
            legend.Position = LegendPos.Right;
            legend.IsReverse = true;
            legend.IsHStack = false;

            chart.AxisChange(); // Fa in modo che il grafico venga ridisegnato
            string percorsoImmagine = "./" + index.ToString() + finefileBmp;

            // Salvo il grafico come immagine e poi lo carico nel pdf
            graphControl.MasterPane.ReSize(graphControl.CreateGraphics(), new RectangleF(0, 0, 2000, 1200));
            Bitmap img = graphControl.GraphPane.GetImage();

            // Se la bitmap esiste già (se ad esempio l'esecuzione precedente non è andata a buon fine e non
            // sono state cancellate correttamente), prima la cancello e poi la ricreo
            if (File.Exists(percorsoImmagine))
            {
                FileInfo fi = new FileInfo(percorsoImmagine);
                fi.IsReadOnly = false;
                File.Delete(percorsoImmagine);
            }

            img.Save(percorsoImmagine);
            // Le bitmap devono essere ReadOnly, almeno finchè non si finisce di confezionare il PDF
            FileInfo fInfo = new FileInfo(percorsoImmagine);
            fInfo.IsReadOnly = true;

            MigraDoc.DocumentObjectModel.Shapes.Image image = new MigraDoc.DocumentObjectModel.Shapes.Image(percorsoImmagine);
            image.Width = Unit.FromCentimeter(28);
            image.Height = Unit.FromCentimeter(15.5);
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Center;

            Row row = table.AddRow();
            row.Cells[0].MergeRight = 1;
            row.Cells[0].Add(image);

            Disclaimer(row);

            if (index != (numStazioni-1)) section.AddPageBreak(); // Metto un'interruzione di pagina ogni volta, tranne che dopo l'ultima pagina
            img.Dispose();
        }

        private void Disclaimer(Row row) 
        {
            // Inserisco il disclaimer, scritto in verticale a destra del grafico
            TextFrame tf = row.Cells[2].AddTextFrame();
            tf.Orientation = TextOrientation.Upward;
            tf.WrapFormat.Style = WrapStyle.None;

            Paragraph disclaimer = new Paragraph();
            disclaimer.Format.Font.Bold = false;
            disclaimer.Format.Font.Size = 6;
            disclaimer.Format.Alignment = ParagraphAlignment.Left;
            disclaimer.AddFormattedText("\"Composizione e rappresentazione dei dati eseguita con modalità automatiche su dati della rete di stazioni\nmeteorologiche fiduciarie della Regione Sardegna gestita dall\'Agenzia per la Protezione dell'Ambiente della Sardegna,\nARPAS, acquisiti in tempo reale e sottoposti ad un processo automatico di validazione di primo livello\"");
            tf.Height = Unit.FromCentimeter(14);
            tf.Width = Unit.FromCentimeter(2);

            tf.Add(disclaimer);        
        
        }


        private double[] riordinaPunti(String[] elencoCampi)
        {
            double[] elencoOrdinato = new double[elencoCampi.Length - indiceH - 2];
            for (int i = 0; i < elencoOrdinato.Length; i++)
                elencoOrdinato[i] = Convert.ToDouble(elencoCampi[i + 1 + indiceH]);

            int xy = 0;
            double[] X = new double[elencoOrdinato.Length / 2];
            double[] Y = new double[elencoOrdinato.Length / 2];
            
            // Copio le coordinate nei vettori X e Y
            for (int i = 0; i < elencoOrdinato.Length; i = i + 2)
            {
                X[xy] = elencoOrdinato[i];
                Y[xy] = elencoOrdinato[i + 1];
                xy++;
            }
            xy = 0;

            // Scorro il vettore X per trovare eventuali punti con stesso valore - andranno ordinati in modo crescente o decrescente a seconda
            // dei punti precedente e successivo, in modo che la linea verticale che verrà disegnata sia coerente con il resto del grafico
            int p = 0, u = 0;
            double yprec = 0, ysucc = 0;
            // p = indice primo punto, u = quanti punti dopo il p-esimo hanno stessa x 
            for (int i = 0; i < X.Length; i++)
            {
                p = i;
                u = 0;
                
                while (((i + 1 + u) < X.Length) && (X[i] == X[i + 1 + u]))
                {
                    u++;
                }
                
                i = i + u;

                // Tutti gli elementi con stessa x vanno ordinati in modo crescente/decrescente (se yprec < ysucc o  il contrario)

                if (p == 0) yprec = Y[0];
                else yprec = Y[p - 1];
                if ((p == (Y.Length - 1)) || ((p + u + 1) >= Y.Length)) ysucc = Y[Y.Length - 1];
                else ysucc = Y[p + u + 1];

                Array.Sort(Y, p, u + 1); // Ordinamento crescente
                if (yprec > ysucc)
                    Array.Reverse(Y, p, u + 1); // Ordinamento decrescente
                
              }
           
            // Ricompongo il vettore elencoOrdinato dove si alternano x e y
            for (int i = 0; i < elencoOrdinato.Length; i = i + 2)
            {
                elencoOrdinato[i] = X[xy];
                elencoOrdinato[i + 1] = Y[xy];
                xy++;
            }
            return elencoOrdinato;
        }

        private void DefineStyles()
        {
            // Da rivedere se c'è qualcosa da modificare negli stili
            Style style = this.document.Styles["Normal"];
            style.Font.Name = "Arial";
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.Font.Size = 9;

        } //DefineStyles()

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
    } // Class
} // Namespace