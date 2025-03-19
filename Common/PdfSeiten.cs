using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

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
        List<string> lehrer = new List<string>();

        QuellDateiName = dateiName;

        using (var pdfDocument = PdfDocument.Open(dateiName))
        {
            int seitenNummer = 1;
            foreach (Page page in pdfDocument.GetPages())
            {
                foreach (var word in page.GetWords())
                {
                    // Prüfe, ob die linke x-Koordinate (BoundingBox.Left) ungefähr 100 ist
                    if (Math.Abs(word.BoundingBox.Left - 100) < 0.1)
                    {
                        lehrer.Add(word.Text); 
                        //Console.WriteLine($"Seite {seitenNummer}: Text: {word.Text}, Position: {word.BoundingBox}");
                    }
                }                
            }
        }

        // Gib die 3 häufigsten Nennungen aus der Liste "lehrer" aus
        var topLehrer = lehrer.GroupBy(x => x)
                      .OrderByDescending(g => g.Count())
                      .Take(3)
                      .Select(g => new { Name = g.Key, Count = g.Count() });

        Console.WriteLine("Die 3 häufigsten Nennungen:");
        foreach (var item in topLehrer)
        {
            Console.WriteLine($"Name: {item.Name}, Häufigkeit: {item.Count}");
        }   
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

    internal void ZähleOffeneKlassenbuchEinträge(object lehrer)
    {
        throw new NotImplementedException();
    }
}