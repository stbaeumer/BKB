using System.Text;
//using PdfReader = iTextSharp.text.pdf.PdfReader;

public class PdfSeiten : List<PdfSeite>
{
    public string QuellDateiName { get; set; }

    public void ZwischenseitenZuordnen()
    {
        if (this.Count(x => x.Student != null) > 0)
        {
            var seitenProSchüler = this.Count / this.Count(x => x.Student != null);

            if (this.Count(x => x.Student != null) > 0)
            {
                for (int z = 1; z < this.Count; z = z + seitenProSchüler)
                {
                    int ersteSeite = z;
                    int letzteSeite = z + seitenProSchüler - 1;
                    var schueler = this.Where(x => x.Seite >= ersteSeite && x.Seite <= letzteSeite)
                        .Select(x => x.Student).FirstOrDefault();

                    if (schueler != null)
                    {
                        for (int ii = z; ii < letzteSeite; ii++)
                        {
                            this[ii].Student = schueler;
                        }
                    }
                }
            }
        }
    }

    public void Read(string dateiName)
    {
        /*
        QuellDateiName = dateiName;

        using (var pdfReader = new PdfReader(dateiName))
        {
            for (int i = 1; i <= pdfReader.NumberOfPages; i++)
            {
                StringBuilder text = new StringBuilder();
                text.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i));
                string inhalt = text.ToString();
                Add(new PdfSeite(i, inhalt, dateiName));
            }

            Global.ZeileSchreiben(dateiName + " - Seiten:", this.Count().ToString(), Global.Farbe.Weiss);
        }*/
    }

    public string GetDatum()
    {
        List<string> datum = new List<string>();

        foreach (var pdfSeite in this)
        {
            var aa = pdfSeite.DatumFinden();

            foreach (var a in aa)
            {
                datum.Add(a);
            }
        }

        return datum.GroupBy(s => s) // Gruppiere die Strings
            .OrderByDescending(g => g.Count()) // Sortiere nach Häufigkeit
            .FirstOrDefault()?.Key; // Nimm die häufigste Gruppe und gib den Schlüssel zurück;
    }

    public string GetArt(List<string> suchmuster)
    {
        List<string> art = new List<string>();

        foreach (var pdfSeite in this)
        {
            var aa = pdfSeite.SuchmusterAnwenden(suchmuster);

            foreach (var a in aa)
            {
                if (!art.Contains(a))
                {
                    art.Add(a);
                }
            }
        }

        if (art.Count == 1)
        {
            return art[0];
        }

        if (art.Count == 0)
        {
            Console.WriteLine("Art nicht erkannt.");
            Console.ReadKey();
        }

        if (art.Count > 1)
        {
            Console.WriteLine("Art nicht eindeutig");
            Console.ReadKey();
        }

        return "";
    }
}