using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

public class PdfDatei
{
    public string DateiName { get; set; }
    public Students Students { get; set; }
    public PdfSeiten PdfSeiten { get; set; }
    public string Art { get; set; }
    public string Datum { get; set; }

    public PdfDatei(string dateiName)
    {
        DateiName = dateiName;
        PdfSeiten = new PdfSeiten();
        Students = new Students();
    }

    public Students GetStudentsMitSeiten(Students students)
    {
        Students studentsMitSeiten = new Students();

        foreach (var pdfSeite in PdfSeiten)
        {
            // Der passende Student zu dieser Seite wird ermittelt.
            Student student = pdfSeite.SeiteZuStudentZuordnen(students);

            if (student != null)
            {
                Student st = studentsMitSeiten.Where(s => s.Vorname == student.Vorname && s.Nachname.ToLower() == student.Nachname.ToLower() && s.Geburtsdatum == student.Geburtsdatum).FirstOrDefault();

                // Wenn es den Student noch nicht in der Liste gibt
                if (st == null)
                {
                    student.PdfSeiten = new PdfSeiten();
                    student.PdfSeiten.Add(pdfSeite);
                    studentsMitSeiten.Add(student);
                } 
            }
            // Wenn auf geraden Seiten kein Student gefunden wurde, ...
            if (pdfSeite.Seite % 2 == 0 && student == null)
            {
                // ... aber auf der Seite zuvor
                if (studentsMitSeiten.Last().PdfSeiten.Last().Seite == pdfSeite.Seite - 1)
                {
                    // und die Seite eine Rückseite ist, erkennbar an der Schulnummer
                    if (pdfSeite.Inhalt.Contains("177659"))
                    {
                        studentsMitSeiten.Last().PdfSeiten.Add(pdfSeite);
                    }    
                }
            }
        }
        return studentsMitSeiten;
    }

    public void SeitenAusQuelldateienLöschen()
    {
        List<int> del = new List<int>();
        foreach (var student in this.Students)
        {
            foreach (var pdfSeite in student.PdfSeiten)
            {
                del.Add(pdfSeite.Seite);
            }
        }

        PdfDocument document = PdfReader.Open(DateiName, PdfDocumentOpenMode.Modify);
        
        foreach (var pdfSeite in PdfSeiten.OrderByDescending(x => x.Seite).Where(x => x.Student != null))
        {
            document.Pages.RemoveAt(pdfSeite.Seite - 1);
        }

        document.Save(DateiName);
    }
}