using System.Dynamic;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using CookComputing.XmlRpc;
using Microsoft.Extensions.Configuration;

namespace Common;

public class Menüeintrag
{
    public Gruppen Gruppen { get; set; }
    public List<string> Beschreibung { get; set; }
    public Dateien Quelldateien { get; set; }
    public Anrechnungen Anrechnungen { get; set; }
    public bool DateienFehlenOderSindLeer { get; set; }
    public string Titel { get; set; }
    public Action<Menüeintrag> Funktion { get; } // Funktion mit Menüeintrag als Parameter
    public Students Students { get; set; }
    public Klassen Klassen { get; set; }

    /// <summary>
    /// Interessierende Klassen
    /// </summary>
    public List<string> IKlassen { get; set; }

    /// <summary>
    /// Interessierende Students
    /// </summary>
    public Students IStudents { get; set; }
    public Datei Zieldatei { get; set; }

    public Menüeintrag(string titel, Anrechnungen anrechnungen, Dateien quelldateien, Students students, Klassen klassen,List<string> beschreibung, Action<Menüeintrag> funktion)
    {
        Titel = titel;
        Anrechnungen = anrechnungen;
        Quelldateien = quelldateien;
        Students = students;
        Klassen = klassen;
        DateienFehlenOderSindLeer = false;
        Beschreibung = beschreibung;
        Funktion = funktion;
        Gruppen = new Gruppen();
        IKlassen = new List<string>();
        IStudents = new Students();
    }

    public Datei Ausführen()
    {        
        Funktion?.Invoke(this);        
        return Zieldatei;
    }

    public Datei LernabschnittsdatenAlt(string zielDateiname, IConfiguration configuration)
    {
        var zieldatei = new Datei(zielDateiname);

        FilterInteressierendeStudentsUndKlassen(configuration);

        var basisDa = Quelldateien.GetMatchingList("basisdaten", IStudents, Klassen);
        if (basisDa == null || !basisDa.Any()) return new Datei(zielDateiname);

        var atlzeug = Quelldateien.GetMatchingList("atlantis-zeugnisse", IStudents, Klassen);
        if (atlzeug == null || !atlzeug.Any()) return new Datei(zielDateiname);

        var klassen = Quelldateien.GetMatchingList("klassen", IStudents, Klassen);
        if (klassen == null || !klassen.Any()) return new Datei(zielDateiname);

        var quit = Konfig("MaxDateiAlter", true, "Maximales Alter der eingelesenen Dateien", Datentyp.Int);
        if (quit) return [];

        foreach (var student in IStudents)
        {
            foreach (var zeugnisdatum in student.GetZeugnisDatums(atlzeug))
            {
                var zeile = student.Nachname + ", " + student.Vorname + ", " + student.Klasse + ", Zeugnisdatum: " +
                            zeugnisdatum.ToShortDateString();

                var recBasis = basisDa
                    .Where(record =>
                    {
                        var dict = (IDictionary<string, object>)record;
                        return dict["Vorname"].ToString() == student.Vorname &&
                               dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                    })
                    .FirstOrDefault();

                if(recBasis == null) continue;

                var recKla = klassen
                    .Where(record =>
                    {
                        var dict = (IDictionary<string, object>)record;
                        return dict["InternBez"].ToString() == student.Klasse;
                    })
                    .FirstOrDefault();

                if (recKla == null) continue;
                
                var dictKla = (IDictionary<string, object>)recKla;                

                var recAtl = atlzeug
                .Where(record =>
                {
                    var dict = record as IDictionary<string, object>;
                    return dict != null &&
                        dict["Field1"]?.ToString()?.Replace("'", "") == student.Nachname &&
                        dict["Field3"]?.ToString()?.Replace("'", "") == student.Geburtsdatum &&
                        dict["Field4"]?.ToString()?.Replace("'", "") == zeugnisdatum.ToShortDateString();
                })
                .FirstOrDefault();

                if (recAtl == null) continue;
                    
                var dictAtl = (IDictionary<string, object>)recAtl;                
                                
                var dictBasis = (IDictionary<string, object>)recBasis;

                var jahr = student.GetJahr(zeugnisdatum);
                zeile += ", Jahr: " + jahr;
                var abschnitt = student.GetAbschnitt(zeugnisdatum);
                zeile += ", Abschnitt: " + abschnitt;
                var jahrgang = student.GetJahrgang(klassen, jahr, zeugnisdatum, recBasis);
                zeile += ", Jg: " + jahrgang;

                dynamic record = new ExpandoObject();
                record.Nachname = student.Nachname;
                record.Vorname = student.Vorname;
                record.Geburtsdatum = student.Geburtsdatum;
                record.Jahr = jahr;
                record.Abschnitt = abschnitt;
                record.Jahrgang = jahrgang;
                record.Klasse = student.Klasse;
                record.Schulgliederung = dictKla["Gliederung"];
                record.OrgForm = dictKla["OrgForm"];
                record.Klassenart = dictKla["Klassenart"];
                record.Fachklasse = "";
                record.Förderschwerpunkt = "";
                record.ZWEIPUNKTLEERZEICHENFörderschwerpunkt = "";
                record.Schwerstbehinderung = student.Schwerstbehinderung;
                record.Wertung = "J";
                record.Wiederholung = "N";
                record.Klassenlehrer = "";
                record.Versetzung = "";
                record.Abschluss = "";
                record.Schwerpunkt = "";
                if (dictAtl != null)
                {
                    record.Konferenzdatum = dictAtl["Field5"].ToString().Replace("'", "");
                }
                else
                {
                    record.Konferenzdatum = string.Empty;
                }
                record.Zeugnisdatum = zeugnisdatum.ToShortDateString();
                record.SummeFehlstd = "";
                record.SummeFehlstdLEERZEICHENunentschuldigt = "";
                record.allgPUNKTLEERZEICHENbildenderLEERZEICHENAbschluss = "";
                record.berufsbezPUNKTLEERZEICHENAbschluss = "";
                record.Zeugnisart = "";
                record.FehlstundenMINUSGrenzwert = "";
                // Ein Jahrgang kleiner als 01 deutet auf eine Laufbahn vor der aktuellen Klasse hin und wird ignoriert.
                if (jahrgang.EndsWith("00")) continue;
                // -- deutet auf einen noch älteren Abschnitt hin
                if (jahrgang.Contains("--") || jahrgang.Contains("00")) continue;
                zieldatei.Add(record);
                Global.ZeileSchreiben(zeile, "ok", ConsoleColor.White, ConsoleColor.Gray);
            }
        }

        return zieldatei;
    }

    /// <summary>
    /// Die IStundets und die IKlassen (List<string>) werden als Eigenschaft des Menüeintrags initialisiert.
    /// </summary>
    public void FilterInteressierendeStudentsUndKlassen(IConfiguration configuration)
    {        
        var klassen = Global.Entschluesseln(configuration["Klassen"]);

        var interessierendeKlassen = new List<string>();
        var interessierendeStudents = new Students();
        //var interessierendeStudents.DateiPfad = this.DateiPfad;

        var linkeSeite = "Sie haben diese Klassen gewählt:";
        var rechteSeite = "keine";
        
        Console.WriteLine("");
        Console.WriteLine("   Bitte die interessierende(n) Klasse(n) (kommasepariert) eingeben.");
        Console.WriteLine(klassen == ""
            ? "   Mit ENTER wählen Sie alle Klassen aus."
            : "   Geben Sie 'alle' ein, um die Auswahl der Klassen nicht einzuschränken.");

        Console.WriteLine(
            "   Sie können auch nur den oder die Anfangsbuchstaben kommaspariert eingeben. Abbruch mit q:");

        Console.WriteLine("");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("      Ihre Auswahl [" + klassen + "] : ");
        Console.ResetColor();

        var x = Console.ReadLine()!.ToLower();

        if (x == "ö")
        {
            Global.OpenCurrentFolder();
        }

        if (x == "x")
        {
            Global.OpenWebseite("https://wiki.svws.nrw.de/mediawiki/index.php?title=Schnittstellenbeschreibung");
        }

        if ((x == "" && klassen == "") || (x.ToLower() == "alle") || (x == "" && klassen == "alle"))
        {
            interessierendeKlassen.AddRange(
                (from s in this.Students where !string.IsNullOrEmpty(s.Klasse) select s.Klasse)
                .Distinct().ToList());
            rechteSeite = "alle Klassen";
            Global.Speichern("Klassen", "alle");
        }
        else
        {
            if (x == "" && klassen != "")
            {
                x = klassen;
            }

            foreach (var klasse in (from s in Students select s.Klasse).Distinct().ToList())
            {
                foreach (var item in x.Trim().Split(','))
                {
                    if (klasse.ToLower().StartsWith(item.ToLower()))
                    {
                        if (!interessierendeKlassen.Select(s => s.ToLower()).Contains(klasse.ToLower()))
                        {
                            if (klasse != "")
                            {
                                interessierendeKlassen.Add(klasse);
                            }
                        }
                    }
                }
            }

            if (interessierendeKlassen.Count > 0)
            {
                Global.Speichern("Klassen", x);
                rechteSeite = string.Join(",", interessierendeKlassen);
            }
            else
            {
                // Wenn q oder ein anderer Buchstabe getippt wurde, dem keine Klasse zugeordnet werden kann, werden 0 SuS zurückgegeben.
                //return new Students();
            }
        }

        IStudents = new Students();
        IStudents.AddRange(from t in Students where interessierendeKlassen.Contains(t.Klasse) select t);
        IKlassen = interessierendeKlassen;
        Console.WriteLine("");
        
        Global.ZeileSchreiben(linkeSeite, rechteSeite, ConsoleColor.White, ConsoleColor.Green);
        linkeSeite = (string.Join(",", interessierendeKlassen.Order()));
        linkeSeite = linkeSeite.Substring(0, Math.Min(Console.WindowWidth / 2, linkeSeite.Length));

        if (interessierendeKlassen.Count > 3)
        {
            linkeSeite = linkeSeite + " (" + interessierendeKlassen.Count().ToString() + " Klassen)";
        }

        if (linkeSeite != "")
        {
            Global.ZeileSchreiben(linkeSeite, IStudents.Count().ToString(), ConsoleColor.Gray, ConsoleColor.White);
        }
    }

    public Datei? LeistungsdatenAlt(string unterordnerUndDateiname)
    {
        var fehlendeFaecher = new List<string>();

        var basisdaten = Quelldateien.GetMatchingList("basisdaten", IStudents, Klassen);
        //var gostdaten = GetMatchingList("gost");
        var atlantiszeugnisse = Quelldateien.GetMatchingList("atlantis-zeugnisse", IStudents, Klassen);
        var klassen = Quelldateien.GetMatchingList("klassen", IStudents, Klassen);
        var faecher = Quelldateien.GetMatchingList("faecher", IStudents, Klassen);

        if (atlantiszeugnisse == null || atlantiszeugnisse.Count == 0 || klassen == null || klassen.Count == 0 || faecher == null || faecher.Count == 0) return [];

        var zieldatei = new Datei(unterordnerUndDateiname);

        foreach (var student in IStudents.Where(x => !x.Klasse.StartsWith("G")))
        {
            foreach (var zeugnisdatum in student.GetZeugnisDatums(atlantiszeugnisse))
            {
                var writeline = student.Nachname + ", " + student.Vorname + ", " + student.Klasse + ", ";

                List<dynamic> recAtl = atlantiszeugnisse
                .Where(record =>
                {
                    var dict = record as IDictionary<string, object>;
                    return dict != null &&
                        dict["Field1"]?.ToString()?.Replace("'", "") == student.Nachname &&
                        dict["Field3"]?.ToString()?.Replace("'", "") == student.Geburtsdatum &&
                        dict["Field4"]?.ToString()?.Replace("'", "") == zeugnisdatum.ToShortDateString() &&
                        !string.IsNullOrEmpty(dict["Field8"].ToString()); // Nur benotete Fächer
                })
                .GroupBy(record => // keine doppelten Einträge 
                {
                    var dict = record as IDictionary<string, object>;
                    return dict != null ? new
                    {
                        Nachname = dict["Field1"]?.ToString()?.Replace("'", ""),
                        Vorname = dict["Field2"]?.ToString()?.Replace("'", ""),
                        Geburtsdatum = dict["Field3"]?.ToString()?.Replace("'", ""),
                        Zeugniskonferenzdatum = dict["Field4"]?.ToString()?.Replace("'", ""),
                        Kurztext = dict["Field9"]?.ToString()?.Replace("'", "")
                    } : null;
                })
                .Select(group => group.First())
                .ToList();

                var recBasis = basisdaten
                    .Where(record =>
                    {
                        var dict = (IDictionary<string, object>)record;
                        return dict["Vorname"].ToString() == student.Vorname &&
                               dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                    })
                    .FirstOrDefault();
                var dictBasis = (IDictionary<string, object>)recBasis;

                var jahr = student.GetJahr(zeugnisdatum);
                writeline += ", " + jahr;
                var abschnitt = student.GetAbschnitt(zeugnisdatum);
                writeline += ", " + abschnitt;
                var jahrgang = student.GetJahrgang(klassen, jahr, zeugnisdatum, dictBasis);
                writeline += ", " + jahrgang;

                foreach (var zeile in recAtl)
                {
                    var dict = (IDictionary<string, object>)zeile;
                    var fach = dict["Field9"].ToString();

                    fach = fach.Replace("  ", " ").Replace("B1", "").Replace("C1", "").Replace("A1", "")
                        .Replace("  ", " ")
                        .Replace("B2", "").Replace("C2", "").Replace("A2", "")
                        .Replace(" GD", " G1").Replace(" GE", " G1").Replace(" GB", " G1").Replace("  ", " ");

                    var kursart = GetKursart(jahrgang, fach);
                    var note = dict["Field8"].ToString();
                    var tendenz = dict["Field10"].ToString();
                    var punkte = dict["Field11"].ToString();
                    var noteOderPunkte = dict["Field13"].ToString();

                    dynamic record = new ExpandoObject();
                    record.Nachname = student.Nachname;
                    record.Vorname = student.Vorname;
                    record.Geburtsdatum = student.Geburtsdatum;
                    record.Jahr = jahr;
                    record.Abschnitt = abschnitt;
                    record.Fach = fach.Replace("'", "").Replace("**", "");
                    record.Fachlehrer = "";
                    record.Kursart = kursart;
                    record.Kurs = "";
                    record.Note = student.GetNote(note, noteOderPunkte, punkte, fach, tendenz);
                    record.Abiturfach = "";
                    record.WochenstdPUNKT = "";
                    record.ExterneSchulnrPUNKT = "";
                    record.Zusatzkraft = "";
                    record.WochenstdPUNKTLEERZEICHENZK = "";
                    record.Jahrgang = jahrgang;
                    record.Jahrgänge = "";
                    record.FehlstdPUNKT = "";
                    record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";

                    // Doppelminus deutet auf einen noch ältern Abschnitt hin
                    if (!jahrgang.Contains("--") && !jahrgang.Contains("00"))
                    {
                        zieldatei.Add(record);
                        var writeline2 = writeline + ", Fach: " + dict["Field6"].ToString() + ", Note:" +
                                         dict["Field8"].ToString();
                        Global.ZeileSchreiben(writeline2, "ok", ConsoleColor.Green, ConsoleColor.White);
                    }
                }
            }
        }

        if (fehlendeFaecher.Count() <= 0) return zieldatei;
        Global.ZeileSchreiben("Es fehlen Fächer müssen in SchILD angelegt werden:",
            fehlendeFaecher.Count().ToString(), ConsoleColor.DarkYellow, ConsoleColor.White);
        Global.ZeileSchreiben("Fächer: ", string.Join(", ", fehlendeFaecher), ConsoleColor.Green, ConsoleColor.White);

        return zieldatei;
    }

    private static string GetKursart(string jahrgang, string fach)
    {
        if (!jahrgang.StartsWith("GY")) return "PUK";
        if (!jahrgang.EndsWith("02") && !jahrgang.EndsWith("03") && !jahrgang.EndsWith("12") &&
            !jahrgang.EndsWith("13")) return "PUK";
        if (!fach.Contains(" L")) return fach.Contains(" G") ? "GKS" : "PUK";
        var linkerTeil = fach.Split(' ')[0].TrimEnd();
        return new List<string>() { "D", "M", "E", "BI" }.Contains(linkerTeil) ? "LK1" : "LK2";
    }

    public Datei Lernabschnittsdaten(string zieldateiname)
    {
       var absencePerStud = Quelldateien.GetMatchingList("absenceperstudent", IStudents, Klassen);
        if (absencePerStud == null || !absencePerStud.Any()) return [];

        var schuelerLernab = Quelldateien.GetMatchingList("lernabschnittsdat", IStudents, Klassen);
        if (schuelerLernab == null || !schuelerLernab.Any()) return [];

        var schuelerBasisd = Quelldateien.GetMatchingList("schuelerbasisdate", IStudents, Klassen);
if (schuelerBasisd == null || !schuelerBasisd.Any()) return [];

        var zielDatei = new Datei(zieldateiname, new Datei(schuelerLernab));

        var quit = Konfig("Abschnitt", true, "Abschnitt eingeben");
        if (quit) return [];
        quit = Konfig("MaxDateiAlter", true, "Maximales Alter der eingelesenen Dateien", Datentyp.Int);
        if (quit) return [];

        switch (Global.Abschnitt)
        {
            case "1":
                quit = Konfig("Halbjahreskonferenzdatum", true, "Halbjahreskonferenzdatum eingeben",
                    Datentyp.DateTime);
                if (quit) return [];
                quit = Konfig("Halbjahreszeugnisdatum", true, "Halbjahreszeugnisdatum eingeben",
                    Datentyp.DateTime);
                if (quit) return [];
                break;
            case "2":
                quit = Konfig("Jahreskonferenzdatum", true, "Jahreskonferenzdatum eingeben",
                    Datentyp.DateTime);
                if (quit) return [];
                quit = Konfig("Jahreszeugnisdatum", true, "Jahreszeugnisdatum eingeben",
                    Datentyp.DateTime);
                if (quit) return [];
                break;
        }

        Konfig("MaximaleAnzahlFehlstundenProTag", true, "Maximale Anzahl zählender Fehlstunden pro Tag", Datentyp.Int);
        Konfig("FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt", true,
            "Anzahl Tage vor der Zeugniskonferenz, an denen Fehlzeiten unberücksichtigt bleiben",
            Datentyp.Int);

        var records = new List<dynamic>();

        try
        {
            foreach (var student in IStudents)
            {
                var dictBasisdaten = schuelerBasisd
                    .Where(recBasis =>
                    {
                        var dictBasis = (IDictionary<string, object>)recBasis;
                        return dictBasis["Nachname"].ToString() == student.Nachname &&
                               dictBasis["Vorname"].ToString() == student.Vorname &&
                               dictBasis["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                               dictBasis["Jahr"].ToString() == Global.AktSj[0] &&
                               dictBasis["Abschnitt"].ToString() == Global.Abschnitt;
                    }).FirstOrDefault() as IDictionary<string, object>;

                if (dictBasisdaten != null)
                {
                    var versetzung = "";
                    var abschluss = "";
                    var klassenlehrer = Klassen.Where(rec => rec.Name == student.Klasse)
                        .Select(rec => rec.Klassenleitungen[0].Kürzel).FirstOrDefault();
                    var jahrgang = string.IsNullOrEmpty(dictBasisdaten!["Jahrgang"].ToString())
                        ? ""
                        : dictBasisdaten["Jahrgang"].ToString();
                    var schulgliederung = string.IsNullOrEmpty(dictBasisdaten["Schulgliederung"].ToString())
                        ? ""
                        : dictBasisdaten["Schulgliederung"].ToString();
                    var orgForm = string.IsNullOrEmpty(dictBasisdaten["OrgForm"].ToString())
                        ? ""
                        : dictBasisdaten["OrgForm"].ToString();
                    var klassenart = string.IsNullOrEmpty(dictBasisdaten["Klassenart"].ToString())
                        ? ""
                        : dictBasisdaten["Klassenart"].ToString();
                    var fachklasse = string.IsNullOrEmpty(dictBasisdaten["Fachklasse"].ToString())
                        ? ""
                        : dictBasisdaten["Fachklasse"].ToString();
                    var zeugnisart = "";
                    var schwerstbehinderung = student.Schwerstbehinderung;
                    var wiederholung = "";

                    var dictLernabschnitt = schuelerLernab
                        .Where(recLern =>
                        {
                            var dictLern = (IDictionary<string, object>)recLern;
                            return dictLern["Nachname"].ToString() == student.Nachname &&
                                   dictLern["Vorname"].ToString() == student.Vorname &&
                                   dictLern["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                   dictLern["Jahr"].ToString() == Global.AktSj[0] &&
                                   dictLern["Abschnitt"].ToString() == Global.Abschnitt;
                        }).FirstOrDefault() as IDictionary<string, object>;

                    var konferenzdatum = Global.Halbjahreskonferenzdatum;
                    var zeugnisdatum = Global.Halbjahreszeugnisdatum;
                    var fehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt =
                        Global.FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt;

                    // Wenn bereits Lernabschnittsdaten existieren, werden die Daten dort entnommen.
                    if (dictLernabschnitt != null)
                    {
                        konferenzdatum = konferenzdatum.Year == 1
                            ? string.IsNullOrEmpty(dictLernabschnitt["Konferenzdatum"].ToString())
                                ? konferenzdatum
                                : Convert.ToDateTime(dictLernabschnitt["Konferenzdatum"].ToString())
                            : konferenzdatum;
                        zeugnisdatum = zeugnisdatum.Year == 1
                            ? string.IsNullOrEmpty(dictLernabschnitt["Zeugnisdatum"].ToString())
                                ? zeugnisdatum
                                : Convert.ToDateTime(dictLernabschnitt["Zeugnisdatum"].ToString())
                            : zeugnisdatum;
                        jahrgang = string.IsNullOrEmpty(jahrgang)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Jahrgang"].ToString())
                                ? jahrgang
                                : dictLernabschnitt["Jahrgang"].ToString()
                            : jahrgang;
                        orgForm = string.IsNullOrEmpty(orgForm)
                            ? string.IsNullOrEmpty(dictLernabschnitt["OrgForm"].ToString())
                                ? orgForm
                                : dictLernabschnitt["OrgForm"].ToString()
                            : orgForm;
                        klassenart = string.IsNullOrEmpty(klassenart)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Klassenart"].ToString())
                                ? klassenart
                                : dictLernabschnitt["Klassenart"].ToString()
                            : klassenart;
                        schulgliederung = string.IsNullOrEmpty(schulgliederung)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Schulgliederung"].ToString())
                                ? schulgliederung
                                : dictLernabschnitt["Schulgliederung"].ToString()
                            : schulgliederung;
                        klassenlehrer = string.IsNullOrEmpty(klassenlehrer)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Klassenlehrer"].ToString())
                                ? klassenlehrer
                                : dictLernabschnitt["Klassenlehrer"].ToString()
                            : klassenlehrer;
                        versetzung = string.IsNullOrEmpty(versetzung)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Versetzung"].ToString())
                                ? versetzung
                                : dictLernabschnitt["Versetzung"].ToString()
                            : versetzung;
                        abschluss = string.IsNullOrEmpty(abschluss)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Abschluss"].ToString())
                                ? abschluss
                                : dictLernabschnitt["Abschluss"].ToString()
                            : abschluss;
                        fachklasse = string.IsNullOrEmpty(fachklasse)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Fachklasse"].ToString())
                                ? fachklasse
                                : dictLernabschnitt["Fachklasse"].ToString()
                            : fachklasse;
                        zeugnisart = string.IsNullOrEmpty(zeugnisart)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Zeugnisart"].ToString())
                                ? zeugnisart
                                : dictLernabschnitt["Zeugnisart"].ToString()
                            : zeugnisart;
                        schwerstbehinderung = string.IsNullOrEmpty(schwerstbehinderung)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Schwerstbehinderung"].ToString())
                                ? schwerstbehinderung
                                : dictLernabschnitt["Schwerstbehinderung"].ToString()
                            : schwerstbehinderung;
                        wiederholung = string.IsNullOrEmpty(wiederholung)
                            ? string.IsNullOrEmpty(dictLernabschnitt["Wiederholung"].ToString())
                                ? wiederholung
                                : dictLernabschnitt["Wiederholung"].ToString()
                            : wiederholung;
                    }

                    dynamic record = new ExpandoObject();
                    record.Nachname = student.Nachname;
                    record.Vorname = student.Vorname;
                    record.Geburtsdatum = student.Geburtsdatum;
                    record.Jahr = Global.AktSj[0];
                    record.Abschnitt = Global.Abschnitt;
                    record.Jahrgang = jahrgang;
                    record.Klasse = student.Klasse;
                    record.Schulgliederung = schulgliederung;
                    record.OrgForm = orgForm;
                    record.Klassenart = klassenart;
                    record.Fachklasse = fachklasse;
                    record.Förderschwerpunkt = "";
                    record.ZWEIPUNKTLEERZEICHENFörderschwerpunkt = "";
                    record.Schwerstbehinderung = schwerstbehinderung;
                    record.Wertung = "J";
                    record.Wiederholung = wiederholung;
                    record.Klassenlehrer = klassenlehrer;
                    record.Versetzung = versetzung;
                    record.Abschluss = abschluss;
                    record.Schwerpunkt = "";
                    record.Konferenzdatum = konferenzdatum.ToShortDateString();
                    record.Zeugnisdatum = zeugnisdatum.ToShortDateString();
                    record.SummeFehlstd = student.GetFehlstd(absencePerStud,
                        fehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt);
                    record.SummeFehlstdUNTERSTRICHunentschuldigt = student.GetUnentFehlstd(absencePerStud,
                        fehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt);
                    record.allgPUNKTMINUSbildenderLEERZEICHENAbschluss = "";
                    record.berufsbezPUNKTLEERZEICHENAbschluss = "";
                    record.Zeugnisart = zeugnisart;
                    record.FehlstundenMINUSGrenzwert = "";
                    record.DatumLEERZEICHENvon = "";
                    record.DatumLEERZEICHENbis = "";
                    records.Add(record);
                }
            }

            zielDatei.AddRange(records);
            return zielDatei;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Console.ReadKey();
        }

        return null;
    }

    public Datei Leistungsdaten(string zieldateiname, IConfiguration configuration, string art = "")
    {
        var zieldatei = new Datei(zieldateiname);
        
        if(art == "Mahnung")
        {
            IStudents = Students;
        }

        var marksPerLs = Quelldateien.GetMatchingList("marksperlesson", IStudents, Klassen);
        if (marksPerLs == null || marksPerLs.Count == 0) return [];

        var expLessons = Quelldateien.GetMatchingList("exportlessons", IStudents, Klassen);
        if (expLessons == null || expLessons.Count == 0) return [];
        
        var stdgroupSs = Quelldateien.GetMatchingList("studentgroupstudents", IStudents, Klassen);
        if (art == "" && (stdgroupSs == null || stdgroupSs.Count == 0)) return [];
        
        List<dynamic> schLeistus = Quelldateien.GetMatchingList("schuelerleistungsdaten", IStudents, Klassen);
        if (art == "" && (schLeistus == null || schLeistus.Count == 0)) return [];

        List<dynamic> schBasisds = Quelldateien.GetMatchingList("schuelerbasisdaten", IStudents, Klassen);
        if (art == "" && (schBasisds == null || schBasisds.Count == 0)) return [];

        var records = new List<dynamic>();

        Global.Konfig("Abschnitt", configuration, "Abschnitt eingeben.");

        if (art == "Mahnung")
        {
            marksPerLs = marksPerLs.Where(rec => 
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Prüfungsart"].ToString().Contains("Mahnung");
            }).ToList();       

            // Reduziere die IStudents-Liste basierend auf den gefilterten marksPerLs
            var x = IStudents.Where(student =>
                marksPerLs.Any(mark =>
                {
                    var dict = (IDictionary<string, object>)mark;
                    return dict["Name"].ToString().Contains(student.Vorname) &&
                        dict["Name"].ToString().Contains(student.Nachname) &&
                        dict["Klasse"].ToString() == student.Klasse;
                })
            ).ToList();

            IStudents.Clear();
            IStudents.AddRange(x);
        }

        foreach (var klasse in IStudents.OrderBy(x => x.Klasse).Select(x => x.Klasse).Distinct())
        {
            Global.DisplayHeader(Global.Header.H3, klasse, Global.Protokollieren.Nein);
            var isFirstRun = true;

            var verschiedeneFaecherDerKlasse = VerschiedeneFaecher(klasse, expLessons);

            foreach (var student in IStudents.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname)
                         .Where(x => x.Klasse == klasse))
            {
                foreach (var fach in verschiedeneFaecherDerKlasse)
                {
                    // Normalerweise gibt es nur einen Unterricht. 
                    var unterrichteMitDiesemFach = GetUnterrichteMitDiesemFach(fach, klasse, expLessons);

                    var dictExp = (IDictionary<string, object>)unterrichteMitDiesemFach[0];

                    var zusatzlehrkraft = "";
                    var zusatzlehrkraftWochenstunden = "";

                    if (!student.UnterrichtIstRelevantFürZeugnisInDiesemAbschnitt(dictExp)) continue;

                    // Wenn dieses Fach mit diesem Lehrer bereits in den records existiert,
                    // dann wird es nicht erneut hinzugefügt.

                    var gibtDasFachMitDemLehrerSchon = records.Any(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return dict["Fach"].ToString() == dictExp["subject"].ToString() &&
                               dict["Fachlehrer"].ToString() == dictExp["teacher"].ToString() &&
                               dict["Vorname"].ToString() == student.Vorname &&
                               dict["Nachname"].ToString() == student.Nachname &&
                               dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                    });

                    if (!gibtDasFachMitDemLehrerSchon)
                    {
                        string jahrgang = student.GetJahrgang(schBasisds);
                        string note = student.GetNote(jahrgang, marksPerLs, dictExp["subject"].ToString()!);
                        string kursart = GetKursart(jahrgang, fach);
                        bool mahnung = student.GetMahnung(marksPerLs, dictExp["subject"].ToString()!);
                       

                        // Die Kursart 
                        var kursartBisher = schLeistus
                            .Where(record =>
                            {
                                var dict = (IDictionary<string, object>)record;
                                return dict["Vorname"].ToString() == student.Vorname &&
                                       dict["Nachname"].ToString() == student.Nachname &&
                                       dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                       dictExp["subject"] != null &&
                                       dict["Fach"].ToString() == dictExp["subject"].ToString();
                            })
                            .Select(record =>
                            {
                                var dict = (IDictionary<string, object>)record;
                                return dict["Kursart"].ToString();
                            })
                            .FirstOrDefault()
                            ?.ToString();

                        if (!string.IsNullOrEmpty(kursartBisher))
                            kursart = kursartBisher;

                        // Klassenunterrichte werden immer hinzugefügt
                        if (dictExp["studentgroup"].ToString() == "")
                        {
                            dynamic record = new ExpandoObject();
                            record.Nachname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Jahr = Global.AktSj[0];
                            record.Abschnitt = Global.Abschnitt;
                            record.Fach = dictExp["subject"].ToString();
                            record.Fachlehrer = dictExp["teacher"].ToString();
                            record.Kursart = kursart;
                            record.Kurs = "";
                            record.Note = note;
                            record.Abiturfach = "";
                            record.WochenstdPUNKT = dictExp["periods"];
                            record.ExterneLEERZEICHENSchulnrPUNKT = "";
                            record.Zusatzkraft = zusatzlehrkraft;
                            record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                            record.Jahrgang = "";
                            record.Jahrgänge = "";
                            record.FehlstdPUNKT = ""; // Fehlzeiten werden über die Abschnittsdaten importiert.
                            record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                            if (art == "Mahnung")
                            {
                                record.Mahnung = "J";
                                record.Sortierung = "";
                                record.Mahndatum = DateTime.Now.ToShortDateString();
                            }
                            if((mahnung && art == "Mahnung") || art != "Mahnung")
                            {
                                records.Add(record);
                            }
                        }
                        else // Bei Kursunterrichten wird geschaut, ob der Schüler den Kurs belegt hat. 
                        {
                            var id = string.IsNullOrEmpty(student.ExterneIdNummer)
                                ? student.IdSchild
                                : student.ExterneIdNummer;
                            var studentZeile = stdgroupSs
                                .Where(record =>
                                {
                                    var dict = (IDictionary<string, object>)record;
                                    return dict["studentId"].ToString() == id &&
                                           dict["studentgroup.name"].ToString() ==
                                           dictExp["studentgroup"].ToString();
                                })
                                .FirstOrDefault();
                            var dictStudentgroup = (IDictionary<string, object>)studentZeile!;

                            if (dictStudentgroup != null)
                            {
                                if (!student.UnterrichtIstRelevantFürZeugnisInDiesemAbschnitt(dictStudentgroup))
                                    continue;
                                dynamic record = new ExpandoObject();
                                record.Nachname = student.Nachname;
                                record.Vorname = student.Vorname;
                                record.Geburtsdatum = student.Geburtsdatum;
                                record.Jahr = Global.AktSj[0];
                                record.Abschnitt = Global.Abschnitt;
                                record.Fach = dictStudentgroup["subject"].ToString();
                                record.Fachlehrer = dictExp["teacher"].ToString();
                                record.Kursart = kursart;
                                record.Kurs = dictStudentgroup["studentgroup.name"].ToString()!.Substring(0,
                                    Math.Min(dictStudentgroup["studentgroup.name"].ToString()!.Length, 20));
                                record.Note = note;
                                record.Abiturfach = "";
                                record.WochenstdPUNKT = dictExp["periods"];
                                record.ExterneLEERZEICHENSchulnrPUNKT = "";
                                record.Zusatzkraft = zusatzlehrkraft;
                                record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                                record.Jahrgang = student.Jahrgang;
                                record.Jahrgänge = "";
                                record.FehlstdPUNKT = "";
                                record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                                if (art == "Mahnung")
                                {   
                                    record.Mahnung = "";
                                    record.Sortierung = "";
                                    record.Mahndatum = DateTime.Now.ToShortDateString();
                                }
                                if((mahnung && art == "Mahnung") || art != "Mahnung")
                                {
                                    records.Add(record);
                                }
                            }
                        }
                    }
                }
            }
        }

        zieldatei.AddRange(records);
        return zieldatei;
    }

    private List<dynamic>? GetUnterrichteMitDiesemFach(string fach, string klasse, List<dynamic>? exportLessons)
    {
        return exportLessons.Where(rec =>
        {
            var dictExp = (IDictionary<string, object>)rec;
            return dictExp["subject"].ToString() == fach && dictExp["klassen"].ToString().Split(' ').Contains(klasse);
        }).OrderByDescending(rec =>
        {
            var dictExp = (IDictionary<string, object>)rec;
            return DateTime.ParseExact(dictExp["endDate"].ToString()!, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        }).ToList();
    }

    public List<string> VerschiedeneFaecher(string klasse, List<dynamic>? exportLessons)
    {
        return exportLessons.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["klassen"].ToString().Split(' ').Contains(klasse);
            })
            .Select(record =>
            {
                var dict = (IDictionary<string, object>)record;
                return dict["subject"].ToString();
            }).Distinct().ToList();
    }

    public Datei Kurse(string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);
        var records = new List<dynamic>();
        var klassen = Quelldateien.GetMatchingList("Klassen", IStudents, Klassen);
        if (!klassen.Any()) return null;
        var exportLessons = Quelldateien.GetMatchingList("Exportlessons", IStudents, Klassen);
        if (!exportLessons.Any()) return null;

        foreach (var recExp in exportLessons)
        {
            var dictExp = (IDictionary<string, object>)recExp;

            if (string.IsNullOrEmpty(dictExp["studentgroup"].ToString())) continue;
            foreach (var klasse in dictExp["klassen"].ToString()!.Split('~'))
            {
                var zeileKlasse = klassen.Where(record =>
                {
                    var dict = (IDictionary<string, object>)record;
                    return dict["InternBez"].ToString() == klasse;
                }).FirstOrDefault();

                if (!IStudents.Select(x => x.Klasse).ToList().Contains(klasse)) continue;
                {
                    dynamic record = new ExpandoObject();
                    record.KursBez = dictExp["studentgroup"].ToString()!.Substring(0,
                        Math.Min(dictExp["studentgroup"].ToString()!.Length, 20));
                    record.Klasse = klasse;
                    record.Jahr = Global.AktSj[0];
                    record.Abschnitt = Global.Abschnitt;
                    record.Jahrgang = zeileKlasse?.Jahrgang;
                    record.Fach = Global.PrüfeAufNullOrEmpty(dictExp, "subject");
                    record.Kursart = GetKursart(record.Jahrgang, dictExp["subject"].ToString());
                    record.WochenstdPUNKT = dictExp["periods"];
                    record.WochenstdPUNKTLEERZEICHENKL = "";
                    record.Kursleiter = dictExp["teacher"];
                    record.Epochenunterricht = "";
                    record.Schulnr = "177659";
                    record.WochenstdPUNKTLEERZEICHENZK = "";
                    record.Zusatzkraft = "";
                    record.WochenstdLEERZEICHENZK = "";
                    record.WeitereLEERZEICHENZusatzkraft = "";
                    zieldatei.AddRange(record);
                }
            }
        }

        return zieldatei;
    }

    public bool Konfig(string parameter, bool immerAnzeigen = false, string beschreibung = "", Datentyp datentyp = Datentyp.String)
    {
var documentsFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        var configPath = Path.Combine(documentsFolderPath, "BKB.json");

        beschreibung = string.IsNullOrEmpty(beschreibung) ? "Bitte " + parameter + " eingeben " : beschreibung;
        var configuration = new ConfigurationBuilder()
            .AddJsonFile(configPath, optional: false, reloadOnChange: true).Build();
        var value = configuration[parameter];
        var property = typeof(Global).GetProperty(parameter, BindingFlags.Public | BindingFlags.Static);

        if (value != "null")
        {
            if (datentyp == Datentyp.Int)
            {
                if (value == "")
                {
                    property?.SetValue(0, Convert.ChangeType(0, property.PropertyType));
                }
                else
                {
                    property?.SetValue(Convert.ToInt32(value),
                        Convert.ChangeType(Convert.ToInt32(value), property.PropertyType));
                }
            }
            else if (datentyp == Datentyp.DateTime)
            {
                if (value == "")
                {
                    property?.SetValue(DateTime.Now.Date, Convert.ChangeType(DateTime.Now.Date, property.PropertyType));
                }
                else
                {
                    property?.SetValue(value, Convert.ChangeType(value, property.PropertyType));
                }
            }
            else
            {
                property?.SetValue(value, Convert.ChangeType(value, property.PropertyType));
            }
        }

        if (property != null && property.CanWrite)
        {
            object? currentValue = property.GetValue(null);
            bool isEmpty = currentValue == null || (currentValue is string str && string.IsNullOrEmpty(str));

            object? convertedValue = null;

            string userInput = currentValue?.ToString();
            bool firstrun = true;

            while (true)
            {
                if (immerAnzeigen)
                {
                    string currentValueStr = !string.IsNullOrEmpty(userInput) ? currentValue.ToString() : "";

                    Console.ForegroundColor = ConsoleColor.Green;

                    Console.Write($"  {beschreibung} [{currentValueStr.Replace(" 00:00:00", "")}] : ");

                    userInput = Console.ReadLine();
                    Console.ResetColor();

                    // Wenn Enter gedrückt wird, bleibt der aktuelle Wert erhalten
                    if (string.IsNullOrEmpty(userInput))
                    {
                        //Console.WriteLine($"  {parameter} bleibt unverändert: {currentValueStr}");
                        return false;
                    }

                    if (userInput == "q")
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("     Sie haben die Verarbeitung abgebrochen.");
                        Console.ResetColor();
                        return false;
                    }

                    if (userInput == "ö")
                    {
                        //todo:
                        Global.OpenCurrentFolder();
                    }

                    if (userInput == "x")
                    {
                        Global.OpenWebseite(
                            "https://wiki.svws.nrw.de/mediawiki/index.php?title=Schnittstellenbeschreibung");
                    }
                }

                Console.ForegroundColor = ConsoleColor.Red;
                // Versuch, den neuen Wert in den richtigen Typ zu konvertieren
                switch (datentyp)
                {
                    case Datentyp.Int:
                        if (int.TryParse(userInput, out int intValue))
                        {
                            convertedValue = intValue;
                        }
                        else
                        {
                            if (!firstrun)
                                Console.WriteLine("   Ungültige Eingabe für Int. Versuchen Sie es erneut.");
                            firstrun = false;
                            immerAnzeigen = false;
                            continue;
                        }

                        break;

                    case Datentyp.Url:
                        if (IsValidHttpUrl(userInput))
                        {
                            convertedValue = userInput;
                        }
                        else
                        {
                            if (!firstrun)
                                Console.WriteLine("   Ungültige Eingabe für einen Link. Versuchen Sie es erneut.");
                            firstrun = false;
                            continue;
                        }

                        break;
                    case Datentyp.DateTime:
                        if (DateTime.TryParse(userInput, out DateTime dateValue) || userInput.Contains("0001"))
                        {
                            convertedValue = dateValue;
                        }
                        else
                        {
                            if (!firstrun)
                                Console.WriteLine("   Ungültige Eingabe für DateTime. Versuchen Sie es erneut.");
                            firstrun = false;
                            continue;
                        }

                        break;

                    case Datentyp.String:
                        convertedValue = userInput;
                        break;
                    case Datentyp.JaNein:
                        if (userInput == "j" || userInput == "J" || userInput == "n" || userInput == "N" ||
                            userInput == "y" || userInput == "Y")
                        {
                            convertedValue = userInput;
                        }
                        else
                        {
                            if (!firstrun)
                                Console.WriteLine(
                                    "   Ungültige EIngabe. Erlaubt sind: j und n. Versuchen Sie es erneut.");
                            firstrun = false;
                            immerAnzeigen = true;
                            continue;
                        }

                        break;
                    case Datentyp.Pfad:
                        if (Path.Exists(userInput))
                        {
                            convertedValue = userInput;
                        }
                        else
                        {
                            if (!firstrun)
                                Console.WriteLine("   Ungültiger Pfad. Versuchen Sie es erneut.");
                            firstrun = false;
                            immerAnzeigen = true;
                            continue;
                        }

                        break;

                    default:
                        Console.WriteLine("   Unbekannter Datentyp.");
                        immerAnzeigen = true;
                        continue;
                }

                Console.ResetColor();
                // Falls der Wert erfolgreich geparst wurde, setzen wir ihn und beenden die Schleife
                if (convertedValue == null) continue;
                Global.Speichern(parameter, userInput);
                property.SetValue(null, convertedValue);
                Console.WriteLine($"  {parameter} auf {convertedValue.ToString().Replace(" 00:00:00", "")} gesetzt.");
                break;
            }
        }
        else
        {
            Console.WriteLine($"Eigenschaft '{parameter}' nicht gefunden oder nicht beschreibbar.");
        }

        return true;
    }

    static bool IsValidHttpUrl(string url)
    {
        return Uri.TryCreate(url, UriKind.Absolute, out Uri? uriResult) &&
               (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
    }


    public Datei? Faecher(string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);

        var records = new List<dynamic>();
        var exportLessons = Quelldateien.GetMatchingList("exportlessons", IStudents, Klassen);
        if (!exportLessons.Any())
        {
            return [];
        }

        var schildFaecher = Quelldateien.GetMatchingList("faecher", IStudents, Klassen);
        if (!schildFaecher.Any())
        {
            return [];
        }

        foreach (var recExp in exportLessons)
        {
            var dictExp = (IDictionary<string, object>)recExp;

            var schildFach = schildFaecher.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["InternKrz"].ToString() == dictExp["subject"].ToString();
            }).FirstOrDefault();

            // Wenn es das Fach in SchILD nicht gibt, ...

            if (schildFach != null) continue;
            {
                // ... wird bei Fächern mit Suffix geprüft, ob es bereits ein Schildfach ohne Suffix gibt.

                var subject = dictExp["subject"].ToString();
                var endetMitZiffer = subject.Length > 0 && char.IsDigit(subject[^1]);

                if (endetMitZiffer)
                {
                    var subjectOhneSuffix = subject.Substring(0, subject.Length - 1);
                    // Die Eigenschaften vom Mutterfach werden übernommen
                    var mutterfach = schildFaecher.Where(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return dict["InternKrz"].ToString() == subjectOhneSuffix;
                    }).FirstOrDefault();

                    // Wenn es ein Mutterfach gibt, wird es mit neuem Namen hinzugefügt
                    if (mutterfach != null)
                    {
                        if (records.Any(x => x.InternKrz == dictExp["subject"].ToString())) continue;
                        var dictMutterfach = (IDictionary<string, object>)mutterfach;
                        dynamic record = new ExpandoObject();
                        record.InternKrz = subject;
                        record.StatistikKrz = dictMutterfach["StatistikKrz"].ToString();
                        record.Bezeichnung = dictMutterfach["Bezeichnung"].ToString();
                        record.BezeichnungZeugnis = dictMutterfach["BezeichnungZeugnis"].ToString();
                        record.BezeichnungÜZeugnis = dictMutterfach["BezeichnungÜZeugnis"].ToString();
                        record.Unterrichtssprache = dictMutterfach["Unterrichtsprache"].ToString();
                        record.SortierungLEERZEICHENS1 = dictMutterfach["Sortierung S1"].ToString();
                        record.SortierungLEERZEICHENS2 = dictMutterfach["Sortierung S2"].ToString();
                        record.Gewichtung = dictMutterfach["Gewichtung"].ToString();

                        var gibtEsSchon = records.Any(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["InternKrz"].ToString() == subject;
                        });

                        if (!gibtEsSchon)
                        {
                            records.Add(record);
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("  Das Fach " + subject +
                                          " nicht gefunden. Es wird in SchILD angelegt. Bitte prüfen!");
                        Console.ResetColor();

                        dynamic record = new ExpandoObject();
                        record.InternKrz = subject;
                        record.StatistikKrz = "FB";
                        record.Bezeichnung = subject;
                        record.BezeichnungZeugnis = "";
                        record.BezeichnungÜZeugnis = "";
                        record.Unterrichtssprache = "";
                        record.SortierungLEERZEICHENS1 = "";
                        record.SortierungLEERZEICHENS2 = "";
                        record.Gewichtung = "";

                        var gibtEsSchon = records.Any(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["InternKrz"].ToString() == subject;
                        });

                        if (!gibtEsSchon)
                        {
                            zieldatei.Add(record);
                        }
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine("  Das Fach " + subject +
                                      " nicht gefunden. Es wird in SchILD angelegt. Bitte prüfen!");
                    Console.ResetColor();

                    dynamic record = new ExpandoObject();
                    record.InternKrz = subject;
                    record.StatistikKrz = "FB";
                    record.Bezeichnung = subject;
                    record.BezeichnungZeugnis = "";
                    record.BezeichnungÜZeugnis = "";
                    record.Unterrichtssprache = "";
                    record.SortierungLEERZEICHENS1 = "";
                    record.SortierungLEERZEICHENS2 = "";
                    record.Gewichtung = "";

                    var gibtEsSchon = records.Any(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return dict["InternKrz"].ToString() == subject;
                    });

                    if (!gibtEsSchon)
                    {
                        records.Add(record);
                    }
                }
            }
        }

        return zieldatei;
    }

    public void LuLAnEintragungDerZeugnisnotenErinnern(Lehrers lehrers)
    {
        var leistungsdaten = Quelldateien.GetMatchingList("leistungsdaten", IStudents, Klassen);
        var betreff = "";
        var adressen = "";
        var anrede = "";
        var lul = new List<string?>();
        var eintaege = new List<string>();

        int i = 1;
        foreach (var leistungsdatum in leistungsdaten)
        {
            var dict = (IDictionary<string, object>)leistungsdatum;

            // Wenn keine Note erteilt wurde ...
            if (dict["Note"].ToString() == "")
            {
                var student = IStudents.FirstOrDefault(x =>
                    x.Vorname == dict["Vorname"].ToString() && x.Nachname == dict["Nachname"].ToString() &&
                    x.Geburtsdatum == dict["Geburtsdatum"].ToString());

                eintaege.Add(i.ToString().PadLeft(4) + ". " + student.Klasse.ToString().PadRight(6) + ", " +
                             (dict["Nachname"] + ", " + dict["Vorname"]).ToString().PadRight(20).Substring(0, 19) +
                             ": " + dict["Fachlehrer"].ToString().PadRight(3) + ": " + dict["Fach"]);
                i++;

                if (!lul.Contains(dict["Fachlehrer"].ToString()))
                {
                    lul.Add(dict["Fachlehrer"].ToString());
                }
            }
        }

        if (eintaege.Count > 0)
        {
            Global.DisplayHeader(Global.Header.H3, "Fehlende Noten:", Global.Protokollieren.Nein);
            foreach (var eintrag in eintaege)
            {
                Console.WriteLine(eintrag);
            }

            foreach (var lehrer in lehrers)
            {
                if (lul.Contains(lehrer.Kürzel))
                {
                    adressen += lehrer.Mail + ",";
                }
            }

            Global.DisplayHeader(Global.Header.H3, "Adressen der Lehrkräfte, die noch eintragen müssen:",
                Global.Protokollieren.Nein);
            Console.WriteLine("   " + adressen.TrimEnd(','));
        }
        else
        {
            Console.WriteLine("  Es fehlen keine Noten. Gut so.");
        }
    }

    public void ChatErzeugen(Lehrers lehrers)
    {
        var recExp = Quelldateien.GetMatchingList("ExportLesson", IStudents, Klassen);
        if (recExp.Count == 0) return;

        var verschiedeneLulKuerzel = recExp
            .Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                var klassenString = dict["klassen"].ToString();
                var klassenListe = klassenString.Split('~'); // Zerlegt den String in eine Liste
                return IKlassen.Any(klasse => klassenListe.Contains(klasse)) &&
                       !string.IsNullOrEmpty(dict["teacher"].ToString());
            }).Select(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["teacher"].ToString();
            }).Distinct();

        lehrers.GetTeamsUrl(verschiedeneLulKuerzel, String.Join(',', IKlassen));
    }

    public Datei? WebuntisOderNetmanCsv(string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);        
        List<dynamic>? webuntisStudents = Quelldateien.GetMatchingList("student_", Students, Klassen);
        if (webuntisStudents == null || webuntisStudents.Count == 0) return [];
        var schuelerZusatzdaten = Quelldateien.GetMatchingList("schuelerzusatzdaten", Students, Klassen);
        if (schuelerZusatzdaten == null || schuelerZusatzdaten.Count == 0) return [];
        var schuelerBasisdaten = Quelldateien.GetMatchingList("schuelerbasisdaten", Students, Klassen);
        if (schuelerBasisdaten == null || schuelerBasisdaten.Count == 0) return [];
        var schuelerErzieher = Quelldateien.GetMatchingList("schuelererzieher", Students, Klassen);
        if (schuelerErzieher == null || schuelerErzieher.Count == 0) return [];
        var schuelerAdressen = Quelldateien.GetMatchingList("schueleradressen", Students, Klassen);
        if (schuelerAdressen == null || schuelerAdressen.Count == 0) return [];
        var lehrkraefte = Quelldateien.GetMatchingList("lehrkraefte", Students, Klassen);
        if (lehrkraefte == null || lehrkraefte.Count == 0) return [];
        var klassen = Quelldateien.GetMatchingList("klassen", Students, Klassen);
        if (klassen == null || klassen.Count == 0) return [];

        // Zuerst werden Änderungen bearbeitet:

        var susMitÄnderung = new List<string>()
        {
            "Folgende Änderungen / Neuanlagen:",
            "Nr".PadRight(5) + "Name".PadRight(46) + "Status Änderung".PadRight(20)
        };

        var i = 1;

        foreach (var rec in webuntisStudents)
        {
            if (rec is not IDictionary<string, object> webuntisStudent) continue;

            var schildStudent = Students
            .Where(x =>
                x.Nachname == webuntisStudent["longName"].ToString() &&
                x.Vorname == webuntisStudent["foreName"].ToString() &&
                x.Geburtsdatum == webuntisStudent["birthDate"].ToString())
            .OrderByDescending(x => int.TryParse(x.IdSchild, out var id) ? id : 0) // IdSchild in int umwandeln, Standardwert 0 bei Fehler
            .FirstOrDefault(); // Gibt den Schüler mit der höchsten IdSchild zurück
                    
            if(schildStudent.Nachname == "Dogani" && schildStudent.Vorname == "Elon"){
                string a = "a";
            }

            if (schildStudent == null) continue;

            var id = string.IsNullOrEmpty(schildStudent.ExterneIdNummer)
                ? schildStudent.IdSchild
                : schildStudent.ExterneIdNummer;

            var schildStudentMeldung = (schildStudent.Nachname + ", " + schildStudent.Vorname + ", " + id + " (" + schildStudent.Klasse + ")").PadRight(45);

            // Wenn er aktiv oder Gast ist, wird seine Klassenzugehörigkeit gecheckt.
            if (new List<string>() { "2", "6" }.Contains(schildStudent.Status))
            {
                if (webuntisStudent["klasse.name"].ToString() != schildStudent.Klasse)
                {
                    susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      " + webuntisStudent["klasse.name"].ToString() + " -> " + schildStudent.Klasse);
                    i++;
                }
            }

            // Wenn der SchildStudent nicht aktiv (2) ist und auch kein Gast (Externer) (6) ist ...
            if (!new List<string>() { "2", "6" }.Contains(schildStudent.Status))
            {                
                // Prüfen, ob ein Austrittsdatum vorhanden ist und ob es in der Vergangenheit liegt
                string exitDateString = webuntisStudent["exitDate"]?.ToString() ?? string.Empty;
                
                if (exitDateString != null && !string.IsNullOrEmpty(exitDateString))
                {
                    DateTime exitDate;
                    bool isValidDate = DateTime.TryParseExact(
                    exitDateString, 
                    "dd.MM.yyyy",  // Das erwartete Datumsformat (bspw. "31.07.2025")
                    CultureInfo.InvariantCulture, 
                    DateTimeStyles.None, 
                    out exitDate);

                    if (isValidDate && exitDate >= DateTime.Now)
                    {
                        // Wenn der Schüler in Schild ein Entlassdatum hat, bekommt er das.

                        schildStudent.GetEntlassdatum(schuelerZusatzdaten);

                        if(!string.IsNullOrEmpty(schildStudent.Entlassdatum))
                        {
                            DateTime entl;
                            bool isValid = DateTime.TryParseExact(schildStudent.Entlassdatum,"dd.MM.yyyy",CultureInfo.InvariantCulture, DateTimeStyles.None, out entl);

                            if(isValid && entl >= DateTime.Now)
                            {
                                susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      Austritt: " + DateTime.Now.ToString("dd.MM.yyyy"));
                                schildStudent.Entlassdatum = DateTime.Now.ToShortDateString();
                            }else
                            {
                                susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      Austritt: " + schildStudent.Entlassdatum);
                                schildStudent.Entlassdatum = DateTime.Now.ToShortDateString();
                            }
                        }
                        i++;
                    }
                }
            }
        }
        
        // Ab hier die Neuanlagen
        // Damit Schüler nicht doppelt angelegt werden, wir zuerst

        var uniqueStudents = Students
            .DistinctBy(s => new { s.Vorname, s.Nachname, s.Geburtsdatum })
            .OrderBy(s => s.Klasse)
            .ThenBy(s => s.Nachname)
            .ThenBy(s => s.Vorname);

        foreach (var studen in uniqueStudents)
        {
            // Es kann sein, dass Schüler nach Abschluss als Gast bleiben. 
            // Es wird angenommen, dass der letzte in der Importliste der aktuelle ist.
            var student = Students.OrderByDescending(x => int.TryParse(x.IdSchild, out var id) ? id : 0)
            .FirstOrDefault(x => x.Nachname == studen.Nachname && x.Vorname == studen.Vorname && x.Geburtsdatum == studen.Geburtsdatum);

            if(student == null) continue;

            if(student.Nachname == "Chernivchan"){
                string a = "a";
            }

            // Wenn der Schüler in Webuntis nicht existiert, ...
            if (!webuntisStudents.Any(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["longName"].ToString() == student.Nachname && dict["foreName"].ToString() == student.Vorname && dict["birthDate"].ToString() == student.Geburtsdatum;
                }))
            {
                // ... und der Schüler in Schild aktiv der Gast ist, wird er angelegt
                if (new List<string>() { "2", "6" }.Contains(student.Status))
                {
                    var id = string.IsNullOrEmpty(student.ExterneIdNummer) ? student.IdSchild : student.ExterneIdNummer;
                    susMitÄnderung.Add(((i + ". ").PadRight(5) + student.Nachname + ", " + student.Vorname + ", " + id + " (" + student.Klasse + ")").PadRight(51) + student.Status + "      Neu: " + student.Klasse);
                    i++;
                }
            }

            var sz = schuelerZusatzdaten
                .Where(rec =>
                {
                    if (rec == null) return false;
                    var dict = (IDictionary<string, object>)rec;
                    return dict != null && dict["Nachname"] != null && dict["Nachname"].ToString() == student.Nachname &&
                           dict["Vorname"].ToString() == student.Vorname &&
                           dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                }).LastOrDefault() as IDictionary<string, object>;


            var sb = schuelerBasisdaten
                ?.Where(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["Nachname"].ToString() == student.Nachname &&
                           dict["Vorname"].ToString() == student.Vorname &&
                           dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                }).LastOrDefault() as IDictionary<string, object>;

            var se = schuelerErzieher
                .Where(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["Nachname"].ToString() == student.Nachname &&
                           dict["Vorname"].ToString() == student.Vorname &&
                           dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                }).LastOrDefault() as IDictionary<string, object>;

            var sa = schuelerAdressen
                .Where(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["Nachname"].ToString() == student.Nachname &&
                           dict["Vorname"].ToString() == student.Vorname &&
                           dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                           dict["Adressart"].ToString() == "Betrieb";
                }).LastOrDefault() as IDictionary<string, object>;

                var klasse = klassen
                .Where(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["InternBez"].ToString() == student.Klasse;
                }).LastOrDefault() as IDictionary<string, object>;

                var klassenleitung = "";
                
                if (klasse != null && klasse.ContainsKey("Klassenlehrer"))
                {
                    var dictklassenleitung = lehrkraefte.Where(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["InternKrz"].ToString() == klasse["Klassenlehrer"].ToString();
                        }).LastOrDefault() as IDictionary<string, object>;

                    klassenleitung = dictklassenleitung["Vorname"] + " " + dictklassenleitung["Nachname"];                                    
                }  
                 

            int alter = -1;
            if (DateTime.TryParseExact(student.Geburtsdatum, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime geburtsdatum))
            {
                alter = DateTime.Now.Year - geburtsdatum.Year;

                // Falls der Geburtstag dieses Jahr noch nicht war, Alter um 1 verringern
                if (DateTime.Now < geburtsdatum.AddYears(alter))
                {
                    alter--;
                }
            }

            dynamic record = new ExpandoObject();

            if (Path.GetFileName(zieldateiname).ToLower().Contains("webuntis"))
            {
                record.Schlüssel = string.IsNullOrEmpty(student.ExterneIdNummer)
                    ? student.IdSchild
                    : student.ExterneIdNummer;
                record.EMINUSMail = student.MailSchulisch;
                record.Familienname = student.Nachname;
                record.Vorname = student.Vorname;
                record.Klasse = student.Klasse;
                record.Kurzname = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");
                record.Geschlecht = sb?["Geschlecht"]?.ToString()?.ToUpper() ?? string.Empty;
                record.Geburtsdatum = student.Geburtsdatum;
                record.Eintrittsdatum = "";
                record.Austrittsdatum = student.Status == "2" || student.Status == "6" ? "31.07." + Global.AktSj[1] : sz?["Entlassdatum"].ToString();
                record.Telefon = sz?["Telefon-Nr."].ToString();
                record.Mobil = "";
                record.Strasse = alter >= 18 ? sb?["Straße"].ToString() : se?["Straße"].ToString();
                record.PLZ = alter >= 18 ? sb?["PLZ"].ToString() : se?["PLZ"].ToString();
                record.Ort = alter >= 18 ? sb?["Ort"].ToString() : se?["Ort"].ToString();
                record.ErzName = alter >= 18 ? "" : se?["Vorname 1.Person"].ToString() + " " + se?["Nachname 1.Person"].ToString();
                record.ErzMobil = alter >= 18 ? "" : "";
                record.ErzTelefon = alter >= 18 ? "" : "";
                record.Volljährig = alter >= 18 ? "1" : "0";
                record.BetriebName = sa == null ? "" : sa["Name1"].ToString();
                record.BetriebStrasse = sa == null ? "" : sa["Straße"].ToString();
                record.BetriebPlz = sa == null ? "" : sa["PLZ"].ToString();
                record.BetriebOrt = sa == null ? "" : sa["Ort"].ToString();
                record.BetriebTelefon = sa == null ? "" : sa["1. Tel.-Nr."].ToString();
                record.O365Identität = student.MailSchulisch;
                record.Benutzername = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");
                zieldatei.Add(record);
            }
            else if (Path.GetFileName(zieldateiname).ToLower().Contains("netman"))
            {
                // Netman
                // ed123456	Dagobert	Eggemann	ed123456@students.berufskolleg-borken.de	E01.07.1992	BZ22A	Stappert, Markus
                record.Schlüssel = string.IsNullOrEmpty(student.ExterneIdNummer)
                    ? student.IdSchild
                    : student.ExterneIdNummer;
                record.Kurzname = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");
                record.Vorname = student.Vorname;
                record.Nachname = student.Nachname;                
                record.Mail = student.MailSchulisch;
                record.Passwort = student.Nachname.Substring(0,1).ToUpper() + student.Geburtsdatum;
                record.Klasse = student.Klasse;
                record.Klassenleitung = klassenleitung;
                
                // todo:
                // In Netman werden die Schüler erst 6 Wochen nach Abschluss ausgebucht.
                if(true)
                {
                    zieldatei.Add(record);
                }
            }
            else if (Path.GetFileNameWithoutExtension(zieldateiname).ToLower().Contains("littera"))
            {
                // Littera
                record.GUID = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");
                record.Familienname = student.Nachname;
                record.Vorname = student.Vorname;
                record.Straße = alter >= 18 ? sb?["Straße"].ToString() : se?["Straße"].ToString();
                record.Postleitzahl = alter >= 18 ? sb?["PLZ"].ToString() : se?["PLZ"].ToString();
                record.Ortsname = alter >= 18 ? sb?["Ort"].ToString() : se?["Ort"].ToString();
                record.Geburtsdatum = student.Geburtsdatum;
                record.EMINUSMail = student.MailSchulisch;
                record.ErzieherLEERZEICHEN1DOPPELPUNKTLEERZEICHENTitel = "";
                record.ErzieherLEERZEICHEN1DOPPELPUNKTLEERZEICHENNachname = alter >= 18 ? "" : se?["Nachname 1.Person"].ToString();
                record.ErzieherLEERZEICHEN1DOPPELPUNKTLEERZEICHENVorname = alter >= 18 ? "" : se?["Vorname 1.Person"].ToString();
                record.ErzieherDOPPELPUNKTLEERZEICHENStraße = "";
                record.ErzieherDOPPELPUNKTLEERZEICHENPostleitzahl = "";
                record.ErzieherDOPPELPUNKTLEERZEICHENOrt = "";
                record.ErzieherLEERZEICHEN2DOPPELPUNKTLEERZEICHENTitel = "";
                record.ErzieherLEERZEICHEN2DOPPELPUNKTLEERZEICHENNachname = "";
                record.ErzieherLEERZEICHEN2DOPPELPUNKTLEERZEICHENVorname = "";
                record.ErzieherDOPPELPUNKTLEERZEICHENStraße = "";
                record.ErzieherDOPPELPUNKTLEERZEICHENPostleitzahl = "";
                record.ErzieherDOPPELPUNKTLEERZEICHENOrt = "";
                record.Geschlecht = sb?["Geschlecht"].ToString().ToUpper() == "M" ? "3": "4"; // (wobei „3“ männlich und „4“ weiblich entspricht)
                record.Klasse = student.Klasse; // Pflichtfeld
                record.InterneLEERZEICHENIDMINUSNummer = string.IsNullOrEmpty(student.ExterneIdNummer)
                    ? student.IdSchild
                    : student.ExterneIdNummer; // Pflichtfeld
                record.Jahrgang = student.Jahrgang; // Pflichtfeld 
                record.ErhältLEERZEICHENBAFög = ""; // Pflichtfeld
                zieldatei.Add(record);
            }
        }

        if (susMitÄnderung.Count() > 2)
            Global.DisplayCenteredBox(susMitÄnderung, 97);

        zieldatei.ZippeBilder("PfadZuAtlantisFotos");

        return zieldatei;
    }

    public Datei NetmanCsv(string zieldateiname, IConfiguration configuration)
    {
        var zieldatei = new Datei(zieldateiname);
        var schuelerBasisdaten = Quelldateien.GetMatchingList("schuelerbasisdaten", IStudents, Klassen);
        var schuelerZusatzdaten = Quelldateien.GetMatchingList("schuelerzusatzdaten", IStudents, Klassen);

        if (schuelerBasisdaten == null || schuelerBasisdaten.Count == 0) return [];

        var andereAlsAktive = schuelerBasisdaten
            .Where(recBasis =>
            {
                var dictBasis = (IDictionary<string, object>)recBasis;
                return dictBasis["Status"].ToString() != "2";
            }).Any();

        if (!andereAlsAktive)
        {
            // Wenn es keine Abgänger in den basisdaten gibt

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("    In den schuelerBasisdaten scheinen nur aktive Schüler enthalten zu sein.");
            Console.WriteLine(
                "    Passen Sie den Export aus Schild so an, dass auch Abgang und Abschluss enthalten ist.");
            Console.ResetColor();
            return null;
        }

        Global.Konfig("AusbuchenNachWievielTagen", configuration,
            "Anzahl der Tage zwischen Abschlusszeugnis und dem Abschalten in O365");

        DateTime datenimportLetztesDatum;
        do
        {
            if (!DateTime.TryParse(Global.DatenimportLetztesDatum.ToShortDateString(), out datenimportLetztesDatum))
            {
                Global.Konfig("DatenimportLetztesDatum", configuration, "Wann hat der letzte Abgleich mit O365 stattgefunden?");
            }
        } while (!DateTime.TryParse(Global.DatenimportLetztesDatum.ToShortDateString(), out datenimportLetztesDatum));

        var records = new List<dynamic>();

        foreach (var recBasisdaten in schuelerBasisdaten.OrderBy(student => student.Klasse)
                     .ThenBy(student => student.Nachname)
                     .ThenBy(student => student.Vorname))
        {
            var dictBasisdaten = (IDictionary<string, object>)recBasisdaten;

            var dictZusatzdaten = schuelerZusatzdaten
                .FirstOrDefault(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["Vorname"].ToString() == dictBasisdaten["Vorname"].ToString() &&
                           dict["Nachname"].ToString() == dictBasisdaten["Nachname"].ToString() &&
                           dict["Geburtsdatum"].ToString() == dictBasisdaten["Geburtsdatum"].ToString();
                }) as IDictionary<string, object>;

            var entlassdatum = dictZusatzdaten != null && dictZusatzdaten.ContainsKey("Entlassdatum") 
                ? DateTime.ParseExact(dictZusatzdaten["Entlassdatum"].ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture) 
                : DateTime.MinValue;
            int ausbuchenNachWievielTagen;
            if (!int.TryParse(Global.AusbuchenNachWievielTagen, out ausbuchenNachWievielTagen))
            {
                ausbuchenNachWievielTagen = 30; // Standardwert setzen (z. B. 30 Tage)
            }


            // Wenn Schüler länger als 60 Tage ausgebucht ist, wird er in O365 ausgebucht.
            if (dictBasisdaten["Status"].ToString() != "2" &&
                entlassdatum.AddDays(ausbuchenNachWievielTagen) > DateTime.Today)
            {
                // Diesen Schüler aus o365 ausbuchen
            }

            if (dictBasisdaten["Status"].ToString() == "2")
            {
                // Schüler neu oder im Bestand
            }

            dynamic record = new ExpandoObject();
            record.Nachname = dictBasisdaten["Nachname"].ToString();
            record.Vorname = dictBasisdaten["Vorname"].ToString();
            record.Mail = dictBasisdaten["MailSchulisch"].ToString();
            record.Klasse = dictBasisdaten["Klasse"].ToString();
            zieldatei.Add(record);


            // "Lfd Nr";"Klasse: Schuljahr";"Klasse";"Schüler: Nachname";"Schüler: Vorname";"Schüler: Ort";"Schüler: Plz";"Schüler: Straße";"Schüler: Geschlecht";"Schüler: Telefon";"Schüler: Barcode";"Schüler: Geburtsdatum";"Klasse: Klassenleiter";"Betrieb: Briefadresse";

            // Datei zippen

            Global.DatenimportLetztesDatum = DateTime.Now;
        }

        return zieldatei;
    }

    public Datei GeevooCsv(string absoluterPfad)
    {
        var zieldatei = new Datei(absoluterPfad);
        var records = new List<dynamic>();

        foreach (var student in IStudents.OrderBy(student => student.Klasse).ThenBy(student => student.Nachname)
                     .ThenBy(student => student.Vorname))
        {
            dynamic record = new ExpandoObject();
            record.Nachname = student.Nachname;
            record.Vorname = student.Vorname;
            record.Mail = student.MailSchulisch;
            record.Klasse = student.Klasse;
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public Datei Teilleistungen(string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);
        var records = new List<dynamic>();

        List<dynamic> exportLessons = Quelldateien.FirstOrDefault(datei =>
                datei.UnterordnerUndDateiname.Contains("exportlessons", StringComparison.CurrentCultureIgnoreCase))!
            .ToList();
        List<dynamic> marksPerLessons = Quelldateien.FirstOrDefault(datei =>
                datei.UnterordnerUndDateiname.Contains("marksperlesson", StringComparison.CurrentCultureIgnoreCase))!
            .ToList();

        try
        {
            // Jeder student ...
            foreach (var student in IStudents)
            {
                // muss mindestens einmal im Halbjahr eine SoLei
                foreach (var recMar in marksPerLessons)
                {
                    var dictMar = (IDictionary<string, object>)recMar;

                    var zeileExportLesson = exportLessons
                        .Where(e =>
                        {
                            var dictExp = (IDictionary<string, object>)e;
                            return dictMar["Fach"].ToString() == dictExp["Fach"].ToString() &&
                                   dictMar["Klasse"].ToString() == dictExp["Klasse"].ToString();
                        })
                        .FirstOrDefault();

                    dynamic record = new ExpandoObject();
                    record.Nachname = student.Nachname;
                    record.Vorname = student.Vorname;
                    record.Geburtsdatum = student.Geburtsdatum;
                    record.Jahr = Global.AktSj[0];
                    record.Abschnitt = Global.Abschnitt;
                    record.Fach = record.Fach;
                    record.Datum = record.Datum;
                    record.Teilleistung = record.Prüfungsart;
                    record.Note = record.Note;
                    record.Bemerkung = record.Bemerkung;
                    record.Lehrkraft = zeileExportLesson?.Teacher;
                    zieldatei.Add(record);
                }
            }

            return zieldatei;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Console.ReadKey();
        }

        return null;
    }

    public Datei KlassenErstellen(string absoluterPfad)
    {
        var zieldatei = new Datei(absoluterPfad);

        var schildKlassen = Quelldateien.GetMatchingList("klassen", Students, Klassen);
        if (schildKlassen.Count == 0) return [];
        var untisKlassen = Quelldateien.GetMatchingList("GPU003", Students, Klassen);
        if (untisKlassen.Count == 0) return [];

        var records = new List<dynamic>();

        foreach (var untisKlasse in untisKlassen)
        {
            var dictUntis = (IDictionary<string, object>)untisKlasse;

            if (dictUntis["Field1"].ToString() == "AGG25C")
            {
                string aa = "";
            }

            var klasseVonDerKopiertWird = dictUntis["Field1"].ToString();

            // Wenn es die Klasse in Schild nicht gibt
            if (!schildKlassen.Any(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["InternBez"].ToString() == dictUntis["Field1"].ToString();
                }))
            {
                klasseVonDerKopiertWird = DecreaseNumberInString(dictUntis["Field1"].ToString());
            }

            // Suche die korrespondierende (Vorgänger-)klasse mit allen Schildeigenschaften
            var schildKlasseVonDerKopiertWird = schildKlassen.FirstOrDefault(zeile =>
            {
                var dict = (IDictionary<string, object>)zeile;
                return dict["InternBez"].ToString() == klasseVonDerKopiertWird;
            });

            if (schildKlasseVonDerKopiertWird != null)
            {
                var s = (IDictionary<string, object>)schildKlasseVonDerKopiertWird;
                // Wenn es diese Klasse in SchILD nicht gibt, wird sie angelegt
                dynamic record = new ExpandoObject();
                record.InternBez = dictUntis["Field1"].ToString();
                record.StatistikBez = dictUntis["Field1"].ToString();
                record.SonstigeBez = "";
                record.Jahrgang = s["Jahrgang"].ToString();
                record.Folgeklasse = dictUntis["Field1"].ToString();
                record.Klassenlehrer = dictUntis.ContainsKey("Field30") && dictUntis["Field30"] != null
                    ? dictUntis["Field30"].ToString().Split(',').FirstOrDefault() ?? ""
                    : "";
                record.OrgForm = s["OrgForm"].ToString();
                record.Klassenart = s["Klassenart"].ToString();
                record.Gliederung = s["Gliederung"].ToString();
                record.Fachklasse = s["Fachklasse"].ToString();
                zieldatei.Add(record);
            }
        }

        return zieldatei;
    }

    private string DecreaseNumberInString(string? input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // Regex sucht eine zweistellige Zahl (\d{2})
        Match match = Regex.Match(input, @"\d{2}");

        if (match.Success)
        {
            int number = int.Parse(match.Value) - 1; // Zahl um 1 verringern
            string newNumber = number.ToString("D2"); // Sicherstellen, dass es zweistellig bleibt

            // Ersetze die Zahl
            string result = input.Replace(match.Value, newNumber);

            // Ersetze den letzten Buchstaben durch 'A'
            result = result.Substring(0, result.Length - 1) + "A";

            return result;
        }

        return input; // Falls keine Zahl gefunden wird, bleibt der String unverändert
    }


    public void Sprechtag(Lehrers lehrers, Raums raums, IConfiguration configuration, string hinweis)
    {
        var dokuwikiZugriff = new DokuwikiZugriff(configuration);

        Global.Konfig("WikiSprechtagKleineAenderung", configuration,
            "Handelt es sich um eine kleine Änderung? Kleine Änderungen erzeugen keine neue Version (j/n)",
            Global.Datentyp.JaNein);

        dokuwikiZugriff.Options = new XmlRpcStruct
        {
            { "sum", "Automatische Aktualisierung" },
            { "minor", Global.WikiSprechtagKleineAenderung } // Kein Minor-Edit
        };

        var content = new List<string>();

        var exportLessons = Quelldateien.GetMatchingList("exportlessons", IStudents, Klassen);
        if (exportLessons == null || !exportLessons.Any()) return;

        Global.Konfig("Sprechtagsdatum", configuration, "Datum des Sprechtags angeben (tt.mm.jjjj)", Global.Datentyp.DateTime);
        //Global.Konfig("wikiSprechtagSeite", true, "Seite eingeben, die manipuliert werden soll.");

        hinweis = hinweis.Replace(" nach der allgemeinen Zeugnisausgabe", ", " + Global.Sprechtagsdatum + ",");

        var alleLehrerImUnterrichtKürzel = exportLessons.Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["teacher"].ToString();
        }).Distinct().ToList();

        var alleLehrerImUnterricht = new Lehrers();

        var vergebeneRäume = new Raums();

        foreach (var lehrer in lehrers.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname))
        {
            if (!(from l in alleLehrerImUnterrichtKürzel where lehrer.Kürzel == l select l).Any()) continue;
            // Wenn Raum und Text2 leer sind, dann wird der Lehrer ignoriert 
            if (lehrer.Raum is null or "" && lehrer.Text2 == "") continue;
            alleLehrerImUnterricht.Add(lehrer);

            var r = (from v in vergebeneRäume where v.Raumnummer == lehrer.Raum select v).FirstOrDefault();

            if (r == null)
            {
                if (lehrer.Raum == null) continue;
                // Wenn der Lehrer außer Haus ist, wird sein Raum freigegeben
                if (!lehrer.Text2.ToLower().Contains("außer haus"))
                {
                    vergebeneRäume.Add(new Raum(lehrer.Raum));
                }
            }
            else
            {
                r.Anzahl++;
            }
        }

        content.Add("====== Sprechtag ======");
        content.Add("");

        content.Add(hinweis);

        var i = 1;
        content.Add("");
        content.Add("<WRAP column 15em>");
        content.Add("");
        content.Add("^Name^Raum^");

        var lehrerProSpalteAufSeite2 = ((alleLehrerImUnterricht.Count - 60) / 3) + 1;

        foreach (var l in alleLehrerImUnterricht.OrderBy(x => x.Nachname))
        {
            var raum = string.IsNullOrEmpty(l.Raum) ? "" : l.Raum;

            // Wenn ein KuK außer Haus ist, wird der Raum aus Untis unterdrückt, bleibt aber in Untis für das nächste Jahr stehen. 
            if (!string.IsNullOrEmpty(l.Text2) && l.Text2.ToLower().Contains("außer haus"))
            {
                raum = "";
            }

            content.Add(
                "|" + (l.Geschlecht == "m" ? "Herr " : "Frau ") + (l.Titel == "" ? "" : l.Titel + " ") +
                l.Nachname + (l.Text2 == "" ? "" : " ((" + l.Text2 + "))") + "|" + raum + "|");

            if (i == 20 || i == 40 || i == 60 || i == 60 + lehrerProSpalteAufSeite2 ||
                i == 60 + lehrerProSpalteAufSeite2 * 2)
            {
                content.Add("</WRAP>");
                content.Add("");

                if (i == 60)
                {
                    content.Add("<WRAP pagebreak>");
                }

                content.Add("<WRAP column 15em>");
                content.Add("");
                content.Add("^Name^Raum^");
            }

            i++;
        }

        content.Add("</WRAP>");

        content.Add(
            "Klassenleitungen finden die Einladung als Kopiervorlage im [[sharepoint>:f:/s/Kollegium2/EjakJvXmitdCkm_iQcqOTLwB-9EWV5uqXE8j3BrRzKQQAw?e=OwxG0N|Sharepoint]].\r\n" +
            Environment.NewLine);

        content.Add("");

        var freieR = raums.OrderBy(x => x.Raumnummer)
            .Where(raum => !(from v in vergebeneRäume where v.Raumnummer == raum.Raumnummer select v).Any()).Aggregate(
                @"Sprechtag: Gewünschte Räume müssen in Untis in den Lehrer-Stammdaten eingetragen werden: ",
                (current, raum) => current + (raum.Raumnummer + " "));

        Global.DisplayCenteredBox([freieR], 97);
        dokuwikiZugriff.PutPage("oeffentlich:sprechtag", string.Join("\n", content));

        Global.OpenWebseite("https://bkb.wiki/oeffentlich:sprechtag");
    }

    public Datei Zusatzdaten(string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);

        var datums = Quelldateien.GetMatchingList("DatumsAusAtlantis", IStudents, Klassen);
        if (!datums.Any())
        {
            return [];
        }

        foreach (var student in IStudents)
        {
            var datumsDiesesSchuelers = datums.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Vorname"].ToString() == student.Vorname &&
                       dict["Nachname"].ToString() == student.Nachname &&
                       DateTime.Parse(dict["Geburtsdatum"].ToString()).ToString("dd.MM.yyyy") ==
                       student.Geburtsdatum.ToString();
            }).LastOrDefault();

            var dictD = (IDictionary<string, object>)datumsDiesesSchuelers;

            if (dictD != null)
            {
                dynamic record = new ExpandoObject();

                // Nachname|Vorname|Geburtsdatum|Namenszusatz|Geburtsname|Geburtsort|Ortsteil|Telefon-Nr.|E-Mail|2. Staatsang.|Externe ID-Nr|Sportbefreiung|Fahrschülerart|Haltestelle|Einschulungsart|Entlassdatum|Entlassjahrgang|Datum Schulwechsel|Bemerkungen
                // 

                record.Nachname =
                    student.Nachname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Vorname =
                    student.Vorname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Geburtsdatum =
                    student.Geburtsdatum; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Namenszusatz = "";
                record.Geburtsname = "";
                record.Geburtsort = "";
                record.Ortsteil = "";
                record.TelefonMINUSNrPUNKT = "";
                record.EMINUSMail = "";
                record.ZWEIPUNKTLEERZEICHENStaatsangPUNKT = "";
                record.ExterneLEERZEICHENIDMINUSNr = "";
                record.Sportbefreiung = "";
                record.Fahrschülerart = "";
                record.Haltestelle = "";
                record.Einschulungsart = "";
                record.Entlassdatum = string.IsNullOrEmpty(dictD["Austrittsdatum"].ToString())
                    ? ""
                    : DateTime.Parse(dictD["Austrittsdatum"].ToString()).ToString("dd.MM.yyyy");
                record.Entlassjahrgang = "";
                record.DatumLEERZEICHENSchulwechsel = "";
                record.Bemerkungen = "";
                record.BKAZVO = "";
                record.BeginnBildungsgang = string.IsNullOrEmpty(dictD["Ausbildungsbeginn"].ToString())
                    ? ""
                    : DateTime.Parse(dictD["Ausbildungsbeginn"].ToString()).ToString("dd.MM.yyyy");
                record.Anmeldedatum = string.IsNullOrEmpty(dictD["Aufnahmedatum"].ToString())
                    ? ""
                    : DateTime.Parse(dictD["Aufnahmedatum"].ToString()).ToString("dd.MM.yyyy");
                record.Bafög = "";
                record.EPMINUSJahre = "";
                record.FaxSLASHMobilnr = "";
                record.Ausweisnummer = "";
                record.schulischeLEERZEICHENEMINUSMail = "";
                record.MasernMINUSImpfnachweis = "";
                zieldatei.Add(record);
            }
            else
            {
                string a = "";
            }
        }

        return zieldatei;
    }

    public Datei Basisdaten(string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);
        var records = new List<dynamic>();
        var datums = Quelldateien.GetMatchingList("DatumsAusAtlantis", IStudents, Klassen);
        var basis = Quelldateien.GetMatchingList("basisdaten", IStudents, Klassen);

        foreach (var student in IStudents)
        {
            var basisDiesesSchuelers = basis.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Vorname"].ToString() == student.Vorname &&
                       dict["Nachname"].ToString() == student.Nachname &&
                       DateTime.Parse(dict["Geburtsdatum"].ToString()).ToString("dd.MM.yyyy") ==
                       student.Geburtsdatum.ToString();
            }).FirstOrDefault();

            var dictB = (IDictionary<string, object>)basisDiesesSchuelers;

            var datumsDiesesSchuelers = datums.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Vorname"].ToString() == student.Vorname &&
                       dict["Nachname"].ToString() == student.Nachname &&
                       DateTime.Parse(dict["Geburtsdatum"].ToString()).ToString("dd.MM.yyyy") ==
                       student.Geburtsdatum.ToString();
            }).ToList();

            if (datumsDiesesSchuelers != null && datumsDiesesSchuelers.Count() > 0)
            {

                // Das älteste Datum zählt.
                var recEintrittsdatum = datumsDiesesSchuelers.OrderBy(dynamic =>
                {
                    var dict = (IDictionary<string, object>)dynamic;
                    return dict["Eintrittsdatum"].ToString();
                }).FirstOrDefault();

                var dictEintrittsdatum = (IDictionary<string, object>)recEintrittsdatum;

                var eintrittsdatum = DateTime.Parse(dictEintrittsdatum["Eintrittsdatum"].ToString())
                    .ToString("dd.MM.yyyy");


                dynamic record = new ExpandoObject();

                // Nachname|Vorname|Geburtsdatum|Geschlecht|Status|PLZ|Ort|Straße|Aussiedler|EINSPUNKTLEERZEICHENStaatsangPUNKT|Konfession|StatistikKrzLEERZEICHENKonfession|Aufnahmedatum|AbmeldedatumLEERZEICHENReligionsunterricht|AnmeldedatumLEERZEICHENReligionsunterricht|SchulpflichtLEERZEICHENerfPUNKT|Reform-Pädagogik|NrPUNKTLEERZEICHENStammschule|Jahr|Abschnitt|Jahrgang|Klasse|Schulgliederung|OrgForm|Klassenart|Fachklasse|NochLEERZEICHENfrei|VerpflichtungLEERZEICHENSprachförderkurs|TeilnahmeLEERZEICHENSprachförderkurs|Einschulungsjahr|ÜbergangsempfPUNKTLEERZEICHENJG5|JahrLEERZEICHENWechselLEERZEICHENS1|1PUNKTLEERZEICHENSchulformLEERZEICHENS1|JahrLEERZEICHENWechselLEERZEICHENS2|Förderschwerpunkt|2PUNKTLEERZEICHENFörderschwerpunkt|Schwerstbehinderung|Autist|LSLEERZEICHENSchulnrPUNKT|LSLEERZEICHENSchulform|Herkunft|LSLEERZEICHENEntlassdatum|LSLEERZEICHENJahrgang|LSLEERZEICHENVersetzung|LSLEERZEICHENReformpädagogik|LSLEERZEICHENGliederung|LSLEERZEICHENFachklasse|LSLEERZEICHENAbschluss|Abschluss|SchulnrPUNKTLEERZEICHENneueLEERZEICHENSchule|Zuzugsjahr|GeburtslandLEERZEICHENSchüler|GeburtslandLEERZEICHENMutter|GeburtslandLEERZEICHENVater|Verkehrssprache|DauerLEERZEICHENKindergartenbesuch|EndeLEERZEICHENEingliederungsphase|EndeLEERZEICHENAnschlussförderung
                // 

                record.Nachname =
                    student.Nachname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Vorname =
                    student.Vorname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Geburtsdatum =
                    student.Geburtsdatum; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Geschlecht = string.IsNullOrEmpty(dictB["Geschlecht"].ToString())
                    ? ""
                    : dictB["Geschlecht"].ToString();
                record.Status = string.IsNullOrEmpty(dictB["Status"].ToString()) ? "" : dictB["Status"].ToString();
                record.PLZ = "";
                record.Ort = "";
                record.Straße = "";
                record.Aussiedler = "";
                record.Staatsang = "";
                record.Konfession = "";
                record.StatistikKrzKonfession = "";
                record.Aufnahmedatum = eintrittsdatum;
                record.AbmeldedatumReligionsunterricht = "";
                record.AnmeldedatumReligionsunterricht = "";
                record.Schulpflichterf = string.IsNullOrEmpty(dictB["Schulpflicht erf."].ToString())
                    ? ""
                    : dictB["Schulpflicht erf."].ToString();
                record.ReformPädagogik = "";
                record.NrStammschule = "";
                record.Jahr = "";
                record.Abschnitt = "";
                record.Jahrgang = "";
                record.Klasse = "";
                record.Schulgliederung = "";
                record.OrgForm = "";
                record.Klassenart = "";
                record.Fachklasse = "";
                record.Nochfrei = "";
                record.VerpflichtungSprachförderkurs = "N";
                record.TeilnahmeSprachförderkurs = "N";
                record.Einschulungsjahr = "";
                record.ÜbergangsempJG5 = "";
                record.JahrWechselS1 = "";
                record.SchulformS1 = "";
                record.JahrWechselS2 = "";
                record.Förderschwerpunkt = "";
                record.Förderschwerpunkt2 = "";
                record.Schwerstbehinderung = string.IsNullOrEmpty(dictB["Schwerstbehinderung"].ToString())
                    ? ""
                    : dictB["Schwerstbehinderung"].ToString();
                record.Autist = string.IsNullOrEmpty(dictB["Autist"].ToString()) ? "" : dictB["Autist"].ToString();
                record.LSSchulnr = "";
                record.LSSchulform = "";
                record.Herkunft = "";
                record.LSEntlassdatum = "";
                record.LSJahrgang = "";
                record.LSVersetzung = "";
                record.LSReformpädagogik = "";
                record.LSGliederung = "";
                record.LSFachklasse = "";
                record.LSAbschluss = "";
                record.Abschluss = "";
                record.SchulnrneueSchule = "";
                record.Zuzugsjahr = "";
                record.GeburtslandSchüler = "";
                record.GeburtslandMutter = "";
                record.GeburtslandVater = "";
                record.Verkehrssprache = "";
                record.DauerKindergartenbesuch = "";
                record.EndeEingliederungsphase = "";
                record.EndeAnschlussförderung = "";
                zieldatei.Add(record);
            }
        }

        return zieldatei;
    }

    public Datei GetFaecher(string zieldateiname)
    {

        var faecher = Quelldateien.GetMatchingList("faecher", IStudents, Klassen);
        var zieldatei = new Datei(zieldateiname);

        var verschiedeneFaecher = faecher.Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Bezeichnung"];
        }).ToList().Distinct();

        foreach (var langname in verschiedeneFaecher)
        {
            var fach = faecher.FirstOrDefault(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Bezeichnung"].ToString() == langname.ToString();
            });

            if (langname == "") continue;
            dynamic record = new ExpandoObject();
            record.name = langname;
            record.Kuerzel = fach["InternKrz"].ToString();
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public Datei GetLehrer(string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);
        var lehrkraefte = Quelldateien.GetMatchingList("lehrkraefte", IStudents, Klassen);
        if (lehrkraefte.Count == 0)
        {
            return [];
        }

        foreach (var lehrer in lehrkraefte)
        {
            var dict = (IDictionary<string, object>)lehrer;

            dynamic record = new ExpandoObject();
            record.Kürzel = dict["InternKrz"].ToString();
            record.Vorname = dict["Vorname"].ToString();
            record.Nachname = dict["Nachname"].ToString();
            record.Name = (dict["Titel"].ToString() == "" ? "" : dict["Titel"] + " ") + dict["Vorname"] + " " +
                          dict["Nachname"];
            record.Mail = dict["E-Mail"].ToString();
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public Datei? Praktikanten(List<string> interessierendeKlassenUndJg, string zieldateiname)
    {

        var records = new List<dynamic>();
        var zieldatei = new Datei(zieldateiname);

        var praktikanten = new List<Student>();

        foreach (var item in interessierendeKlassenUndJg)
        {
            praktikanten.AddRange((from s in Students
                where s.Klasse.StartsWith(item.Split(',')[0])                
                select s).ToList());
        }

        foreach (var praktikant in praktikanten)
        {
            if (praktikant == null) continue;
            dynamic record = new ExpandoObject();
            record.Name = praktikant.Nachname + ", " + praktikant.Vorname;
            record.Klasse = praktikant.Klasse;
            record.Jahrgang = praktikant.Jahrgang;
            record.Betrieb = "";
            record.Betreuung = "";
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public Datei? KlassenAnlegen(string zieldateiname)
    {
        var std = Students;
        var kl = Klassen;

        var klassen = Quelldateien.GetMatchingList("klassen", Students, new Klassen());

        if (klassen.Count == 0)
        {
            return [];
        }

        var records = new List<dynamic>();
        
        var zieldatei = new Datei(zieldateiname);

        foreach (var klasse in klassen.OrderBy(x =>
                 {
                     var dictKlasse = (Dictionary<string, dynamic>)x;
                     return x["InternBez"].ToString();
                 }))
        {
            var dictKlasse = (Dictionary<string, dynamic>)klasse;

            dynamic record = new ExpandoObject();
            record.Name = dictKlasse["InternBez"].ToString();
            record.Klassenleitung = dictKlasse["Klassenlehrer"].ToString();
            record.Klassensprecher = "";
            record.Klassensprecher2 = "";
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public void Schulpflichtüberwachung()
    {
        var schuelerMitAbwesenheitenUndMaßnahmen = GetMaßnahmenUndFehlzeiten(
            [
                "Ordnungsmaßnahme",
                "Anhörung",
                "Bußgeldverfahren",
                "Attestpflicht",
                "Mahnung",
                "Familienkasse",
                "Versäumnisanzeige",
                "Suspendierung",
                "ausschluss"
            ]
        );

        schuelerMitAbwesenheitenUndMaßnahmen.SchulpflichtüberwachungTxt(
            "ImportNachWiki/schulpflichtueberwachung.txt",
            10, // Schonfrist: So viele Tage hat die Klassenleitung Zeit offene Stunden
            // zu bearbeiten, bevor eine Warnung ausgelöst wird.
            20, // Nach so vielen unent. Stunden ohne Maßnahme wird eine Warnung ausgelöst.
            30, // Nach so vielen Tagen verjähren unentschuldigte Fehlstunden für Unbescholtene.
            90, // Nach so vielen Tagen verjähren unentschuldigte Fehlstunden für SuS mit Maßnahme
            Klassen,
            Quelldateien
        );
    }

    public Students GetMaßnahmenUndFehlzeiten(List<string> maßnahmenString)
    {
        var sMitAbwesenheiten = new Students();

        var schuelerZusatzdaten = Quelldateien.GetMatchingList("schuelerzusatzdaten", IStudents, Klassen);
            //.Where(datei => datei != null && !string.IsNullOrEmpty(datei!.UnterordnerUndDateiname))
            //.FirstOrDefault(datei => datei!.UnterordnerUndDateiname.Contains("SchuelerZusatzdaten"));
        
        foreach (Student student in Students)
        {
            var id = student.GetId(schuelerZusatzdaten);
            student.GetMaßnahmen(Quelldateien, maßnahmenString);
            student.GetAbwesenheiten(Quelldateien, id);

            if (student.Abwesenheiten.Count != 0)
                sMitAbwesenheiten.Add(student);
        }

        return sMitAbwesenheiten;
    }

    public Datei GetGruppen(string zieldateiname, Anrechnungen anrechnungen, Lehrers lehrers)
    {
        var zieldatei = new Datei(zieldateiname);
        var exportlessons = Quelldateien.GetMatchingList("exportlesson", IStudents, Klassen);
        if (exportlessons == null || exportlessons.Count == 0) return [];

        Gruppen = new Gruppen();        
        Gruppen.AddRange(new Gruppen().GetBildungsgaenge(exportlessons, Klassen, anrechnungen, lehrers));
        Gruppen.AddRange(new Gruppen().GetSchulformen(exportlessons, Klassen, anrechnungen, lehrers));
        Gruppen.Add(new Gruppe().Get(exportlessons, Klassen, anrechnungen, lehrers,
            "versetzung:blaue_briefe",
            new List<string>() { "BS", "HBG", "HBT", "HBW", "FS" },
            new List<int>() { 1 }));
        Gruppen.Add(new Gruppe().Get(exportlessons, Klassen, anrechnungen, lehrers,
            "termine:fhr:start",
            new List<string>() { "BS", "HBG", "HBT", "HBW", "FS", "FM" },
            new List<int>() { 2 }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:deutsch_kommunikation",
            new List<string>() { "D", "D FU", "D1", "D2", "D G1", "D G2", "D L1", "D L2", "D L", "DL", "DL1", "DL2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:englisch",
            new List<string>() { "E", "E FU", "E1", "E2", "E G1", "E G2", "E L1", "E L2", "E L", "EL", "EL1", "EL2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:religionslehre",
            new List<string>() { "KR", "KR FU", "KR1", "KR2", "KR G1", "KR G2", "ER", "ER G1" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:mathematik_physik",
            new List<string>() { "M", "M FU", "M1", "M2", "M G1", "M G2", "M L1", "M L2", "M L", "ML", "ML1", "ML2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:politik_gesellschaftslehre",
            new List<string>() { "PK", "PK FU", "PK1", "PK2", "GG G1", "GG G2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:wirtschaftslehre_in_nicht_kaufm_klassen",
            new List<string>() { "WL", "WBL" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:sport",
            new List<string>() { "SP", "SP G1", "SP G2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:biologie",
            new List<string>() { "BI", "Bi", "Bi FU", "Bi1", "Bi G1", "Bi G2", "BI G1", "BI L1" }));
        Gruppen.Add(new Gruppe().GetKollegium(exportlessons, Klassen, anrechnungen, lehrers,
            ":kollegium:start"));
        Gruppen.Add(new Gruppe().GetLehrerinnen(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:lehrerinnen"));
        Gruppen.Add(new Gruppe().GetRefs(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:referendar_innen"));
        Gruppen.Add(new Gruppe().GetKlassenleitungen(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:klassenleitungen"));

        Gruppen.Add(new Gruppe().GetBildungsgangleitungen(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:bildungsgangleitungen"));
        Gruppen.Add(new Gruppe().GetByWikilink(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:schulleitung:erweiterte:start"));        
        
        foreach (var gruppe in Gruppen)
        {            
            zieldatei.Add(gruppe.Record);
        }
        return zieldatei;
    }


    public Datei MailadressenErgänzen(string zieldateiname)
    {
        var absoluterPfad = Path.Combine(Global.PfadSchilddateien, zieldateiname);

        Konfig("MailDomain", true, "Mail-Domain der Schüler*innen angeben. Bsp: '@students.berufskolleg-borken.de'");

        var zieldatei = new Datei(absoluterPfad);
        var schuelerZusatzdaten = Quelldateien.GetMatchingList("schuelerzusatzdaten", Students, Klassen);
        if (schuelerZusatzdaten.Count == 0) return [];

        // Die Schild-ID muss aus dem SchildSchuelerExport ausgelesen werden
        var students = new Students(@"SchildSchuelerExport", "*.txt");
        if (Students.Count == 0)
        {
            return [];
        }

        foreach (var recZusatz in schuelerZusatzdaten)
        {
            var dictB = (IDictionary<string, object>)recZusatz;
            var vorname = dictB["Vorname"].ToString();
            var nachname = dictB["Nachname"].ToString();
            var geburtsdatum = dictB["Geburtsdatum"].ToString();

            var verschiedeneSuSmitDiesenDaten = schuelerZusatzdaten.Where(rec =>
            {
                var x = (IDictionary<string, object>)rec;
                return vorname == x["Vorname"].ToString() && nachname == x["Nachname"].ToString() &&
                       geburtsdatum == x["Geburtsdatum"].ToString();
            }).ToList();

            var verschiedenestudnts = students.Where(rec =>
                rec.Geburtsdatum == geburtsdatum && rec.Nachname == nachname && rec.Vorname == vorname).ToList();


            for (int i = 0; i < verschiedeneSuSmitDiesenDaten.Count; i++)
            {
                var x = verschiedeneSuSmitDiesenDaten[i];
            }


            var Zusatz = (IDictionary<string, object>)recZusatz;


            var student = students.Where(x =>
                x.Nachname == Zusatz["Nachname"].ToString() && x.Vorname == Zusatz["Vorname"].ToString() &&
                x.Geburtsdatum == Zusatz["Geburtsdatum"].ToString());

            dynamic record = new ExpandoObject();

            record.Nachname = Zusatz["Nachname"].ToString();
            record.Vorname = Zusatz["Vorname"].ToString();
            record.Geburtsdatum = Zusatz["Geburtsdatum"].ToString();
            record.Namenszusatz = Zusatz["Namenszusatz"].ToString();
            record.Geburtsname = Zusatz["Geburtsname"].ToString();
            record.Geburtsort = Zusatz["Geburtsort"].ToString();
            record.Ortsteil = Zusatz["Ortsteil"].ToString();
            record.TelefonMINUSNrPUNKT = Zusatz["Telefon-Nr."].ToString();
            record.EMINUSMail = Zusatz["E-Mail"].ToString();
            record.ZWEIPUNKTLEERZEICHENStaatsangPUNKT = Zusatz["2. Staatsang."].ToString();
            record.ExterneLEERZEICHENIDMINUSNr = Zusatz["Externe ID-Nr"].ToString();
            record.Sportbefreiung = Zusatz["Sportbefreiung"].ToString();
            record.Fahrschülerart = Zusatz["Fahrschülerart"].ToString();
            record.Haltestelle = Zusatz["Haltestelle"].ToString();
            record.Einschulungsart = Zusatz["Einschulungsart"].ToString();
            record.Entlassdatum = Zusatz["Entlassdatum"].ToString();
            record.Entlassjahrgang = Zusatz["Entlassjahrgang"].ToString();
            record.DatumLEERZEICHENSchulwechsel = Zusatz["Datum Schulwechsel"].ToString();
            record.Bemerkungen = Zusatz["Bemerkungen"].ToString();
            record.BKAZVO = Zusatz["BKAZVO"].ToString();
            record.BeginnBildungsgang = Zusatz["BeginnBildungsgang"].ToString();
            record.Anmeldedatum = Zusatz["Anmeldedatum"].ToString();
            record.Bafög = Zusatz["Bafög"].ToString();
            record.EPMINUSJahre = Zusatz["EP-Jahre"].ToString();
            record.FaxSCHRÄGSTRICHMobilnr = Zusatz["FaxSCHRÄGSTRICHMobilnr"].ToString();
            record.Ausweisnummer = Zusatz["Ausweisnummer"].ToString();
            record.schulischeLEERZEICHENEMINUSMail = "";
            throw new Exception();
            record.MasernMINUSImpfnachweis = Zusatz["Masern-Impfnachweis"].ToString();
            record.bisherigeLEERZEICHENID = "";
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public enum Datentyp
    {
        String,
        Int,
        DateTime,
        Pfad,
        Url,
        JaNein
    }

    public Datei Kalender2Wiki(string kalender, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname + ".csv");
        var kalenderRec = Quelldateien.GetMatchingList(kalender, Students, Klassen);
        if (kalenderRec?.Count != 0)
        {
            var records = new List<dynamic>();

            var sortedRecords = kalenderRec?
                .Where(rec =>
                {
                    var beginnString = (string)((IDictionary<string, object>)rec)["Beginn"];
                    var kategorienString = (string)((IDictionary<string, object>)rec)["Kategorien"] ?? "";

                    return beginnString.Split(" ").Length > 0 
                    && !string.IsNullOrEmpty(kategorienString)
                    && DateTime.ParseExact(beginnString.Split(" ")[1], "dd.MM.yyyy", new CultureInfo("de-DE")) >=
                           new DateTime(Convert.ToInt32(Global.AktSj[0]), 07, 31); // keine alten SJ
                })
                .OrderBy(rec =>
                {
                    var beginnString = (string)((IDictionary<string, object>)rec)["Beginn"];
                    return DateTime.ParseExact(beginnString.Substring(3, beginnString.Length - 3), "dd.MM.yyyy HH:mm",
                        new CultureInfo("de-DE"));
                })
                .ToList();

            if (sortedRecords != null)
            {
                foreach (var rec in sortedRecords)
                {
                    var dict = (IDictionary<string, object>)rec;

                    var beginnString = (string)((IDictionary<string, object>)rec)["Beginn"];
                    var endeString = (string)((IDictionary<string, object>)rec)["Ende"];
                    var beginnDatum = DateTime.ParseExact(beginnString.Substring(3, beginnString.Length - 3),
                        "dd.MM.yyyy HH:mm", new CultureInfo("de-DE"));
                    var endeDatum = DateTime.ParseExact(endeString.Substring(3, endeString.Length - 3), "dd.MM.yyyy HH:mm",
                        new CultureInfo("de-DE"));
                    var dat = beginnDatum.ToString("ddd dd.MM.yyyy", new CultureInfo("de-DE"));
                    var zeit = "";

                    // Wenn zwischen beginn und ende exakt 24 Stunden oder ein Vielfaches von 24 liegen, dann ist das Ereignis ganztägig
                    bool ganztaegig = (endeDatum - beginnDatum).TotalHours % 24 == 0;

                    // Bei mehrtägiges, ganztägigen Ereignissen muss das Endedatum um einen Tag nach vorne geschoben werden

                    if ((endeDatum - beginnDatum).TotalHours >= 24 && endeDatum.Hour == 0 && endeDatum.Minute == 0 &&
                        endeDatum.Second == 0)
                    {
                        endeDatum = endeDatum.AddDays(-1);
                    }

                    if (beginnDatum.Hour != 0)
                    {
                        zeit = ", " + beginnDatum.ToShortTimeString();

                        if (endeDatum.Hour != 0)
                        {
                            zeit += " - " + endeDatum.ToShortTimeString();
                        }

                        zeit += " Uhr";
                    }

                    if (ganztaegig && beginnDatum.Date != endeDatum.Date)
                    {
                        dat += " - " + endeDatum.ToString("ddd dd.MM.yyyy", new CultureInfo("de-DE"));
                    }

                    var sj = "vergangene";

                    if (new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1) < beginnDatum &&
                        beginnDatum < new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31))
                    {
                        sj = "aktuelles";
                    }
                    if (beginnDatum > new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31))
                    {
                        sj = "kommendes";
                    }

                    dynamic record = new ExpandoObject();
                    record.Betreff = dict["Betreff"].ToString()!.Trim();
                    record.Seite = dict["Kategorien"].ToString().Split(';')[0];
                    record.Hinweise = "";
                    record.Datum = dat + zeit;
                    record.Kategorien = GetKategorien(dict["Kategorien"].ToString());
                    record.Verantwortlich = "";
                    record.Ort = dict["Ort"].ToString()!.Trim();
                    record.Ressourcen = dict["Ressourcen"].ToString()!.Trim();
                    record.SJ = sj;

                    if (!zieldatei.AbsoluterPfad.Contains("kollegium"))
                    {
                    }
                    else
                    {
                        record.Links = "";
                    }

                    zieldatei.Add(record);
                }
            }
            return zieldatei;
        }

        return [];
    }

    private string GetKategorien(string? toString)
    {
        var kategorien = toString?.Split(';').Aggregate("", (current, str) => current + (str.Trim() + ","));

        return kategorien.TrimEnd(',');
    }
}