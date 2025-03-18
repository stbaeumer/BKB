using System.Text.RegularExpressions;

public partial class PdfSeite
{
    public Student Student { get; set; }
    public string DateiName { get; set; }
    public string Inhalt { get; set; }
    public string Datum { get; set; }
    public int Seite { get; set; }

    public PdfSeite(int seite, string inhalt, string dateiName)
    {
        Seite = seite;
        Inhalt = inhalt;
        DateiName = dateiName;
    }

    public Student SchuelerZuSeiteZuordnen(Students students)
    {
        var studs = new List<Student>();

        foreach (var student in students.Where(student => Inhalt.Contains(student.Nachname, StringComparison.OrdinalIgnoreCase) &&
                                                          Inhalt.Contains(student.Vorname, StringComparison.OrdinalIgnoreCase) &&
                                                          Inhalt.Contains(student.Geburtsdatum, StringComparison.OrdinalIgnoreCase)))
        {
            if (!(studs.Where(st =>
                    st.Nachname == student.Nachname && st.Vorname == student.Vorname &&
                    st.Geburtsdatum == student.Geburtsdatum)).Any())
            {
                studs.Add(student);
            }

            // Suche nach einem weiteren Datum (neben dem Geburtsdatum). Das Datum wird der Eigenschaft Zeugnisdatum zugewiesen
            var dateMatches = Regex.Matches(Inhalt, @"\b\d{2}\.\d{2}\.\d{4}\b");
            foreach (Match dateMatch in dateMatches)
            {
                if (dateMatch.Value.Equals(student.Geburtsdatum)) continue;
                Datum = dateMatch.Value == null ? "" : dateMatch.Value;
                break;
            }

            dateMatches = Regex.Matches(Inhalt, @"Borken,\s(\d{2}\.\d{2}\.\d{4})");
            foreach (Match dateMatch in dateMatches)
            {
                if (dateMatch.Value.Equals(student.Geburtsdatum)) continue;
                try
                {
                    Zeugnisdatum = dateMatch.Groups[1].Value;
                }
                catch
                {
                    // ignored
                }

                break;
            }
        }

        if (studs.Count == 1) return studs[0];
        if (studs.Count <= 1) return null!;
        Console.WriteLine("Mehrere Studs gefunden");
        Console.ReadKey();
        return null!;
    }

    public List<string> SuchmusterAnwenden(List<string> muster)
    {
        List<string> art = new List<string>();

        foreach (var a in from m in muster select Regex.Match(Inhalt, m, RegexOptions.IgnoreCase) into match where match.Success select match.Value == null ? "" : match.Value into a where !art.Contains(a) select a)
        {
            art.Add(a);
        }

        return art;
    }

    public string Zeugnisdatum { get; set; }

    public Student SeiteZuStudentZuordnen(Students students)
    {
        var studs = new List<Student>();

        foreach (var student in students)
        {
            if (student.Vorname == "Mauritz")
            {
                var aa = "";
            }

            if (!Inhalt.Contains(student.Nachname!, StringComparison.OrdinalIgnoreCase) ||
                !Inhalt.Contains(student.Vorname!, StringComparison.OrdinalIgnoreCase) ||
                !Inhalt.Contains(student.Geburtsdatum!, StringComparison.OrdinalIgnoreCase)) continue;
            if (!(studs.Where(st =>
                    st.Nachname == student.Nachname && st.Vorname == student.Vorname &&
                    st.Geburtsdatum == student.Geburtsdatum)).Any())
            {
                studs.Add(student);
            }
        }

        if (studs.Count == 1) return studs[0];
        if (studs.Count <= 1) return null!;
        Console.WriteLine("Mehrere Studs gefunden");
        Console.ReadKey();
        return null!;
    }

    public IEnumerable<string> DatumFinden()
    {
        var datum = new List<string>();

        const string muster = @"\bBorken,?\s(0[1-9]|[12][0-9]|3[01])\.(0[1-9]|1[0-2])\.(\d{4})\b";

        // Wenn Zeugnis als Wort oder Namensbestandteil in der Seite vorkommt, dann gib das Wort zurück.
        var match = Regex.Match(Inhalt, muster, RegexOptions.IgnoreCase);

        if (!match.Success) return datum;
        var a = false ? "" : match.Value;

        if (datum.Contains(a)) return datum;
        const string pattern = @"\b\d{2}\.\d{2}\.\d{4}\b";

        var datumMatch = MyRegex().Match(a);
        datum.Add(datumMatch.Value);

        return datum;
    }

    [GeneratedRegex(@"\b\d{2}\.\d{2}\.\d{4}\b")]
    private static partial Regex MyRegex();
}