using System.Dynamic;
using System.Globalization;
using System.Reflection;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using System.Text.RegularExpressions;
using Common;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using Microsoft.Extensions.Configuration;

public class Students : List<Student>
{
    private Students InteressierendeStudents { get; set; }
    private string DateiPfad { get; set; }
    private DateTime Erstelldatum { get; }
    private DateTime[] SchildVersionExpected { get; set; }

    public Students()
    {
    }

    public Students(Dateien quelldateien)
    {
        var schildschuelerexport = quelldateien
        .Where(x => x?.AbsoluterPfad?.Contains("SchildSchuelerExport") == true)
        .OrderByDescending(x => x?.AbsoluterPfad != null ? File.GetCreationTime(x.AbsoluterPfad) : DateTime.MinValue)
        .FirstOrDefault();

        if (schildschuelerexport == null) return;
        
        foreach (var rec in schildschuelerexport)
        {
            var dict = (IDictionary<string, object>)rec;
            var student = new Student();
                student.Vorname = dict["Vorname"].ToString();
                student.Nachname = dict["Nachname"].ToString();                
                student.Geburtsdatum = dict["Geburtsdatum"].ToString();
                student.Klasse = dict["Klasse"].ToString();
                student.IdSchild = dict["Interne ID-Nummer"].ToString();
                student.ExterneIdNummer = dict["Externe ID-Nummer"].ToString();
                student.MailSchulisch = student.GenerateMail();               
                student.Status = dict["Status"].ToString();
                Add(student);
        }
        
        Global.ZeileSchreiben("Schüler*innen mit Status", this.Count.ToString(), Global.DefaultBackground,ConsoleColor.Blue);

       // Global.InverteUeberschriftZeichnen("Schüler*innen mit Status:", ConsoleColor.Blue,ConsoleColor.Black);
        
        foreach (var status in this.Select(x => x.Status).Distinct())
        {
            int anzahl = this.Count(x => x.Status == status);

            var statusAusgeschrieben = status switch
            {
                "2" => "2: aktiv",
                "3" => "3:",
                "6" => "6: Extern/Gast",
                "8" => "8: Abschluss",
                "9" => "9: Abgänger"
            };
            Global.ZeileSchreiben( "Schüler*innen .......................... Status " + statusAusgeschrieben + ":", anzahl.ToString(),ConsoleColor.Blue, Global.DefaultBackground);
        }
        if (this.Select(x => x.Status).Distinct().Count() == 1)
        {
            Console.WriteLine("    Es scheinen nur aktive Schüler exportiert worden zu sein. Bitte auch alle SuS exportieren.");
            Console.ReadKey();
        }
    }

    public Students(string dateiName, string dateiendung, string delimiter = "|")
    {
        DateiPfad = Global.CheckFile(dateiName, dateiendung);

        if (DateiPfad == null)
        {
            return;
        }

        Erstelldatum = File.GetCreationTime(DateiPfad);

        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HeaderValidated = null,
            MissingFieldFound = null,
            HasHeaderRecord = true,
            Delimiter = delimiter
        };

        using (var reader = new StreamReader(DateiPfad))
        using (var csv = new CsvReader(reader, config))
        {
            csv.Context.RegisterClassMap<SchuelersMap>();
            var records = csv.GetRecords<Student>();

            this.AddRange(records);
            foreach (var student in this)
            {
                student.MailSchulisch = student.GenerateMail();

                if (student.AktuellerAbschnitt == "")
                {
                    student.AktuellerAbschnitt = Global.Abschnitt;
                }

                if (student.AktuellesHalbjahr == "")
                {
                    student.AktuellesHalbjahr = Global.AktSj[0];
                }
            }
        }

        Global.ZeileSchreiben( (DateiPfad + " ").PadRight(90, '.') + " " + Erstelldatum, this.Count().ToString(),
            ConsoleColor.Yellow,ConsoleColor.Gray);
    }

    public bool Keine()
    {
        if (this.Count == 0)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("");
            Console.WriteLine("  Es wurde kein einziger Schüler ausgewählt.");
            Console.WriteLine("  Erneute Auswahl mit ENTER.");
            Console.ReadKey();
            Console.ForegroundColor = ConsoleColor.White;
            return true;
        }

        return false;
    }

    private void ErgänzeFehlendeEigenschaften(Student vS, Student nS)
    {
        foreach (PropertyInfo property in typeof(Student).GetProperties())
        {
            if (property.PropertyType == typeof(string))
            {
                string vSvalue = (string)property.GetValue(vS);

                if (vSvalue.ToLower().Contains("externe"))
                {
                    string aa = "";
                }

                string? nSvalue = (string)property.GetValue(nS)!;

                if (string.IsNullOrEmpty(vSvalue))
                {
                    if (!string.IsNullOrEmpty(nSvalue))
                    {
                        property.SetValue(vS, nSvalue);
                    }
                    else
                    {
                        property.SetValue(vS, "");
                    }
                }
            }
        }
    }

    public void SchulpflichtüberwachungTxt(
        string datei,
        int schonfrist,
        int warnungAbAnzahl,
        int verjaehrungUnbescholtene,
        int nachSovielenTagenVerjährenFehlzeitenBeiMaßnahme,
        Klassen klasses,
        List<Datei> dateien)
    {
        var tempdatei = Path.Combine(Path.GetTempPath(), Path.GetFileName(datei));

        File.WriteAllText(tempdatei,
            "====== Schulpflichtüberwachung ======" + Environment.NewLine, Encoding.UTF8);

        var zeilen = new List<string>();
        zeilen.Add(Environment.NewLine);

        zeilen.Add(@"**Hallo Klassenleitung,**" + Environment.NewLine);
        zeilen.Add(@"" + Environment.NewLine);
        zeilen.Add(
            @"Du wurdest von Teams hierher verlinkt, weil bei der automatisierten, wöchentlichen Durchsicht der Fehlzeiten eine mögliche Schulpflichtverletzung in Deiner Klasse aufgepoppt ist. Können wir Dir Arbeit abnehmen?" +
            Environment.NewLine);
        zeilen.Add(@"" + Environment.NewLine);
        zeilen.Add(@"**Fragen & Antworten**" + Environment.NewLine);
        zeilen.Add(@"" + Environment.NewLine);
        zeilen.Add(
            @"  * :?: Was ist das Ziel dieser Seite? :!: Kritische Fälle erkennen, Reaktionszeiten verkürzen, Klassenleitungen Arbeit abnehmen, SuS signalisieren, dass wir hinschauen." +
            Environment.NewLine);

        zeilen.Add(
            @"  * :?: Wie oft soll ich mahnen? :!: Nach der Mahnung folgt i.d.R. die Teilkonferenz oder das Bußgeldverfahren. Wenn die letzte Mahnung sehr lange her ist, kommt eine weitere Mahnung in Betracht. " +
            Environment.NewLine);

        zeilen.Add(
            @"  * :?: Was, wenn die Zahlen nicht stimmen? :!: Dann gerne melden bei [[chat>stefan.baeumer|Stefan Bäumer]]." +
            Environment.NewLine);

        zeilen.Add(
            @"  * :?: Muss ich eine irgendwem eine Rückmeldung zu den Fällen in meiner Klasse geben? :!: Nein. Eine Rückmeldung ist nicht notwendig. Wer Fragen hat, kann sich natürlich immer melden: [[chat>stefan.baeumer|Stefan Bäumer]]." +
            Environment.NewLine);
        zeilen.Add(@"" + Environment.NewLine);
        zeilen.Add(@"" + Environment.NewLine);
        zeilen.Add(@"" + Environment.NewLine);
        zeilen.Add(@"" + Environment.NewLine);
        zeilen.Add(@"===== Tabelle Schulpflichtüberwachung KW " + Global.Kalenderwoche + "=====" + Environment.NewLine);
        zeilen.Add(@"" + Environment.NewLine);

        zeilen.Add("<searchtable>" + Environment.NewLine);
        zeilen.Add("^  Klasse  ^  Klassenleitung  ^  Name  ^  Alter am 1.Schultag im SJ " + Global.AktSj[0] + "/" +
                   Global.AktSj[1] +
                   "  ^  bisherige Maßnahmen  ^  Aussage  ^Womit können wir Arbeit abnehmen?  ^" + Environment.NewLine);

        string teamsChatLink =
            "chats>sina.milewski@berufskolleg-borken.de,stefan.gantefort@berufskolleg-borken.de,ursula.moritz@berufskolleg-borken.de,";
        var mailliste =
            "mailto:sina.milewski@berufskolleg-borken.de;stefan.gantefort@berufskolleg-borken.de;ursula.moritz@berufskolleg-borken.de;";

        List<dynamic> schuelerZusatzdaten = dateien
            .Where(rec => !string.IsNullOrEmpty(rec.UnterordnerUndDateiname))
            .FirstOrDefault(rec => rec.UnterordnerUndDateiname.Contains("SchuelerZusatzdaten"))!.ToList();

        foreach (var kl in (from k in this.OrderBy(x => x.Klasse) select k.Klasse).Distinct().ToList())
        {
            var klassenleitungen = (from k in klasses where k.Name == kl select k.Klassenleitungen[0]).ToList();

            foreach (var student in this.OrderBy(x => x.Nachname))
            {
                string id = student.GetId(schuelerZusatzdaten);

                if (student.Klasse == kl)
                {
                    var name = student.Vorname!.Substring(0, 2) + "." + student.Nachname!.Substring(0, 2);

                    // Geburtsdatum der Person
                    DateTime geburtsdatum = DateTime.Parse(student.Geburtsdatum!);

                    // Datum, an dem das Alter berechnet werden soll
                    DateTime ersterSchultag = new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1);

                    int alter = ersterSchultag.Year - geburtsdatum.Year;

                    // Prüfen, ob der Geburtstag nach dem 1. August liegt, um das Alter korrekt anzupassen
                    if (geburtsdatum > ersterSchultag.AddYears(-alter))
                    {
                        alter--;
                    }

                    student.GetJüngsteMaßnahmeInDiesemSj();

                    // Für SuS ohne bisherige Maßnahme: Alle Fehlstunden, die noch nicht verjährt sind.


                    /*
                     * |------------|-------------------|-----------------------|-------------------|
                     *            Fehlzeit1           Fehlzeit2              Fehlzeit3            jetzt
                     *            21.8.                24.8.                   27.8.               28.8.
                     *            6 Stunden            5 Stunden               4 Stunden
                     *
                     *                      |<----------------------------------------------------->|
                     *                            Verjährung oder Zeit seit Maßnahme 6 Tage
                     *
                     *                                               |<---------------------------->|
                     *                                                     Schonfrist für KL
                     *                                                     zur Behandlung
                     *                                                     von Fehlzeiten
                     *                                                     3 Tage
                     *
                     *                       |-----------------------|
                     *                        Da die Fehlzeit2 innerhalb der Verjährung
                     *                        aber vor der Schonfrist liegt, werden
                     *                        alle Fehlzeiten (auch die in der Schonfrist) gewarnt.
                     */

                    student.GetF2(verjaehrungUnbescholtene, schonfrist);
                    student.GetF3(verjaehrungUnbescholtene, schonfrist);
                    student.GetF2PlusF3(verjaehrungUnbescholtene, schonfrist);
                    student.GetF2M(verjaehrungUnbescholtene, schonfrist);
                    student.GetF2MplusF3();

                    var aussage = "";
                    var mahnung = "";
                    var mahnungWikiLink = "";
                    var attestpflicht = "";
                    var attestpflichtWikiLink = "";
                    var teilkonferenz = "";
                    var bußgeldverfahren = "";

                    var anzahlMassnahmenInDiesemSj = student.Massnahmen.Where(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return DateTime.Parse(dict["Datum"].ToString()!) >
                               new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1);
                    }).Count();

                    // Wenn es noch keine Maßnahme in diesem SJ gab, ...
                    if (anzahlMassnahmenInDiesemSj == 0)
                    {
                        // ... und wenn es eine F2 gibt ...
                        if (student.F2 > 0)
                        {
                            if (student.F2PlusF3 > warnungAbAnzahl)
                            {
                                // ... dann werden F2 und F3 angemahnt.

                                aussage += student.F2PlusF3 + " unent. Fehlst. in den letzten " +
                                           verjaehrungUnbescholtene + " Tagen. ";
                                mahnung = student.GetUrl("Mahnungen");
                                mahnungWikiLink = student.GetWikiLink("Mahnung", student.F2PlusF3);
                                attestpflicht = student.GetUrl("Attestpflicht");
                                attestpflichtWikiLink = student.GetWikiLink("Attestpflicht", student.F2PlusF3);
                            }
                        }
                    }

                    var schonMassnahmen = student.Massnahmen.Where(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return DateTime.Parse(dict["Datum"].ToString()!) >
                               new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1);
                    }).Count();

                    // Wenn es schon Maßnahmen gab, ...
                    if (schonMassnahmen > 0)
                    {
                        // ... und wenn es eine F2M gibt ...
                        if (student.F2M > 0)
                        {
                            // ... dann werden F2M und F3 angemahnt.

                            var dictS = (IDictionary<string, object>)student.JuengsteMassnahmeInDiesemSj;

                            aussage += student.F2MplusF3 + " unent./offene Fehlstd. seit " +
                                       dictS["Vermerkart"].ToString() + "(" +
                                       dictS["Datum"].ToString() + ").";

                            if (dictS["Vermerkart"].ToString() == "Mahnung")
                            {
                                if (alter < 18)
                                {
                                    bußgeldverfahren =
                                        @"\\ [[eskalationsstufen_erzieherische_einwirkung_ordnungsmassnahmen:bussgeldverfahren:start|Bußgeldverfahren]]";
                                }
                                else
                                {
                                    teilkonferenz =
                                        @"\\ [[eskalationsstufen_erzieherische_einwirkung_ordnungsmassnahmen:bussgeldverfahren:start|Teilkonferenz]]";
                                }
                            }
                        }
                    }

                    student.MaßnahmenAlsWikiLinkAufzählung = student.GetMaßnahmenAlsWikiLinkAufzählung();

                    if (aussage.Length > 0)
                    {
                        var klassenleitungenString = "";

                        foreach (var k in klassenleitungen)
                        {
                            if (!klassenleitungenString.Contains(k + ","))
                            {
                                klassenleitungenString += k.Kürzel + ",";
                            }

                            if (!mailliste.Contains(k.Mail))
                            {
                                mailliste += k.Mail + ";";
                            }

                            if (!teamsChatLink.Contains(k.Mail))
                            {
                                teamsChatLink += k.Mail + ",";
                            }
                        }

                        zeilen.Add("|" + student.Klasse.PadRight(10) + "|" +
                                   klassenleitungenString.TrimEnd(',').PadRight(16) + "  |" + name.PadRight(8) + "|" +
                                   alter + "|" +
                                   student.MaßnahmenAlsWikiLinkAufzählung + "  |" + aussage +
                                   "  |[[:eskalationsstufen_erzieherische_einwirkung_ordnungsmassnahmen|Erz.Einwirkung]] " +
                                   attestpflichtWikiLink + " " + mahnungWikiLink + " " + bußgeldverfahren + " " +
                                   teilkonferenz + "|" + Environment.NewLine);
                    }
                }
            }
        }

        zeilen.Add("</searchtable>" + Environment.NewLine);

        teamsChatLink = teamsChatLink.TrimEnd(',') + "&topicName=Schulpflichtüberwachung KW " + Global.Kalenderwoche +
                        "&message=Bitte beachten: https://bkb.wiki/schulpflichtueberwachung";

        foreach (var zeile in zeilen)
        {
            var z = zeile.Replace("Teams", @"[[" + teamsChatLink + @"|Teams]]");
            File.AppendAllText(tempdatei, z, Encoding.UTF8);
        }

        Global.Dateischreiben(datei);
    }

    public List<dynamic> Reliabmelder()
    {
        UTF8Encoding utf8NoBom = new UTF8Encoding(false);
        var records = new List<dynamic>();
        return records;
    }


    public List<dynamic>? SchuelerFehlstunden(Dateien dateien)
    {
        var records = new List<dynamic>();

        foreach (var student in this.OrderBy(datei => datei.Klasse).ThenBy(zeile => zeile.Nachname).ToList())
        {
            List<dynamic> absencePerStudents =
                dateien.FirstOrDefault(datei => datei.UnterordnerUndDateiname.ToLower().Contains("absenceperstudent"))!
                    .ToList();

            foreach (var recAbs in absencePerStudents)
            {
                var dictAbs = (IDictionary<string, object>)recAbs;

                if (dictAbs["Schüler*innen"].ToString()!.Contains(student.Nachname!) &&
                    dictAbs["Schüler*innen"].ToString()!.Contains(student.Vorname!))
                {
                    if (dictAbs["Datum"] != null)
                    {
                        int fehlstd = string.IsNullOrEmpty(dictAbs["Fehlstd."].ToString())
                            ? 0
                            : Convert.ToInt32(dictAbs["Fehlstd."].ToString());
                        var entschuldigt = "N";
                        if (dictAbs["Status"] != null)
                        {
                            if (dictAbs["Status"].ToString() == "entsch.")
                            {
                                entschuldigt = "J";
                            }
                        }

                        if (fehlstd > 0)
                        {
                            dynamic record = new ExpandoObject();
                            record.Nachname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Klasse = student.Klasse;
                            record.Jahr = Global.AktSj[0];
                            record.Abschnitt = Global.Abschnitt;
                            record.Datum = dictAbs["Datum"].ToString();
                            record.Fach = "";
                            record.von_StdPUNKT = "";
                            record.bis_StdPUNKT = "";
                            record.FehlstdPUNKT = fehlstd;
                            record.Entschuldigt = entschuldigt;
                            record.Lehrkraft = "";
                            records.Add(record);
                        }
                    }
                }
            }
        }

        return records;
    }

    public void Fotos()
    {
        var pfad = Path.Combine(Global.PfadExportdateien, "Fotos-Input");

        if (!Directory.Exists(pfad))
        {
            Directory.CreateDirectory(pfad);
            Console.WriteLine($"   Der Ordner '{pfad}' wurde erstellt.");
        }

        Console.WriteLine("   Hier müssen die Fotos liegen: " + pfad);

        do
        {
            Console.WriteLine("");
            Console.Write("   Klasse eingeben: ");
            var klasse = "";
            do
            {
                klasse = Console.ReadLine();
            } while (!this.Any(x => x.Klasse == klasse.ToUpper()));

            Console.WriteLine("");
            List<Student> sortierteKlasse = this
                .Where(x => x.Klasse == klasse.ToUpper())
                .OrderBy(x => x.Nachname)
                .ThenBy(x => x.Vorname).ToList();
            Console.WriteLine("   Jetzt die Schüler*innen in dieser Reihenfolge fotografieren:");
            var z = 1;
            foreach (var student in sortierteKlasse)
            {
                Console.WriteLine("    " + z.ToString().PadLeft(2) + ". " + student.Nachname + ", " + student.Vorname);
                z++;
            }

            Console.WriteLine("");
            Console.WriteLine("   Liegen die " + sortierteKlasse.Count() + " Fotos der Klasse " + klasse.ToUpper() +
                              " im Ordner '" + pfad + "'? Dann ENTER");

            var x = Console.ReadKey();

            if (x.Key != ConsoleKey.Enter)
            {
                break;
            }

            // Alle jpg-Dateien im Ordner finden
            var jpgDateien = Directory.GetFiles(pfad, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => file.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                               file.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            // Anzahl der Dateien ausgeben
            if (sortierteKlasse.Count() == jpgDateien.Length)
            {
                Console.WriteLine($"   Anzahl der JPG-Bilder im Ordner '{pfad}'");
                Console.WriteLine($"   stimmt mit der Anzahl der SuS in der Klasse überein: {jpgDateien.Length}");

                var klassenpfad = Path.Combine(Global.PfadExportdateien, "Fotos", klasse.ToUpper());
                // Ordner "Klasse" erstellen, falls er nicht existiert
                if (!Directory.Exists(klassenpfad))
                {
                    Directory.CreateDirectory(klassenpfad);
                    Console.WriteLine($"   Der Ordner '{klassenpfad}' wurde erstellt.");
                }

                // Laufende Nummer und Dateiverarbeitung
                for (int i = 0; i < jpgDateien.Length; i++)
                {
                    string bildPfad = jpgDateien[i];
                    FileInfo bildInfo = new FileInfo(bildPfad);

                    // Bildinformationen ausgeben
                    Console.WriteLine($"    {i + 1}. {bildInfo.Name} - Erstellt am: {bildInfo.CreationTime}");

                    // Umbenennung vorbereiten
                    if (i < sortierteKlasse.Count())
                    {
                        string neuerDateiname =
                            $"{sortierteKlasse[i].Nachname}_{sortierteKlasse[i].Vorname}_{sortierteKlasse[i].IdSchild}.jpg";
                        string neuerPfad = Path.Combine(klassenpfad, neuerDateiname);

                        // Datei verschieben und umbenennen
                        File.Move(bildPfad, neuerPfad);

                        Console.WriteLine("     verschoben nach" + neuerPfad);
                    }
                    else
                    {
                        Console.WriteLine(
                            $"    Keine weiteren Studenten in der Liste verfügbar für Datei: {bildInfo.Name}");
                    }

                }
            }
            else
            {
                Console.WriteLine(
                    $"   Anzahl der JPG-Bilder im Ordner '{pfad}' stimmt NICHT der Anzahl der SuS in der Klasse überein.");
            }
        } while (true);
    }

    private sealed class SchuelersMap : ClassMap<Student>
    {
        public SchuelersMap()
        {
            Map(m => m.Vorname).Name(" Vorname", "Vorname", "foreName", "Allg. Adresse: Name2");
            Map(m => m.Aktjahrgang).Name("9 Aktjahrgang");
            Map(m => m.AktuellerAbschnitt).Name("Aktueller Abschnitt");
            Map(m => m.AktuellesHalbjahr).Name("Aktuelles Halbjahr");
            Map(m => m.Aufnahmedatum).Name("22 Aufnahmedatum", "Aufnahmedatum", "entryDate");
            Map(m => m.Adressmerkmal).Name("68 Adressmerkmal");
            Map(m => m.Adressart).Name("Allg. Adresse: Adressart");
            Map(m => m.Anrede).Name("Anrede");
            Map(m => m.Asv).Name("ASV");
            Map(m => m.AufnehmendeSchuleName).Name("Aufnehmende Schule: Name");
            Map(m => m.AufnehmendeSchuleOrt).Name("Aufnehmende Schule: Ort");
            Map(m => m.AufnehmendeSchulePlz).Name("Aufnehmende Schule: PLZ");
            Map(m => m.AufnehmendeSchuleSchulnr).Name("Aufnehmende Schule: Schulnr.");
            Map(m => m.AufnehmendeSchuleStraße).Name("Aufnehmende Schule: Straße");
            Map(m => m.Aussiedler).Name("Aussiedler");
            Map(m => m.Ausweisnummer).Name("Ausweisnummer");
            Map(m => m.Ausbildungsort).Name("24 Ausbildort", "25 Betriebsort");
            Map(m => m.Austritt).Name(" Austritt", "45 EntlDatum", "Austrittsdatum", "Entlassdatum", "exitDate");
            Map(m => m.AbschlussartAnEigenerSchuleStatistikKürzel)
                .Name("Abschlussart an eigener Schule (Statistik-Kürzel)");
            Map(m => m.Bemerkungen).Name("Allg. Adresse: Bemerkungen", "Bemerkungen");
            Map(m => m.BeginnDerBildungsganges).Name("Beginn d. Bildungsganges");
            Map(m => m.BerufsschulpflichtErfüllt).Name("Berufsschulpflicht erfüllt");
            Map(m => m.BesMerkmal).Name("bes. Merkmal");
            Map(m => m.BleibtAnSchule).Name("Bleibt an Schule?");
            Map(m => m.Briefanrede).Name("Briefanrede");
            Map(m => m.Berufsabschluss).Name("65 Berufsabschluss");
            Map(m => m.Berufswechsel).Name("Schüler Berufswechsel");
            Map(m => m.Betreuung).Name("61 Betreuung");
            Map(m => m.BetreuerAbteilung).Name("Allg. Adresse: Betreuer Abteilung");
            Map(m => m.BetreuerAnrede).Name("Allg. Adresse: Betreuer Anrede");
            Map(m => m.BetreuerEmail).Name("Allg. Adresse: Betreuer E-Mail");
            Map(m => m.BetreuerName).Name("Allg. Adresse: Betreuer Name");
            Map(m => m.BetreuerTelefon).Name("Allg. Adresse: Betreuer Telefon");
            Map(m => m.BetreuerTitel).Name("Allg. Adresse: Betreuer Titel");
            Map(m => m.BetreuerVorname).Name("Allg. Adresse: Betreuer Vorname");
            Map(m => m.Betreuungslehrer).Name("Allg. Adresse: Betreuungslehrer");
            Map(m => m.BetreuungslehrerAnrede).Name("Allg. Adresse: Betreuungslehrer Anrede");
            Map(m => m.Bezugsjahr).Name("1 Bezugsjahr");
            Map(m => m.Bkazvo).Name("62 BKAZVO");
            Map(m => m.Branche).Name("Allg. Adresse: Branche");
            Map(m => m.Bundesland).Name("Allg. Adresse: Bundesland", "Bundesland");
            Map(m => m.Einschulungsart).Name("58 Einschulungsart");
            Map(m => m.ElternteilZugezogen).Name("56 Elternteil zugezogen");
            Map(m => m.Entlassjahrgang).Name("Entlassjahrgang");
            Map(m => m.Erzieher1Anrede).Name("Erzieher 1: Anrede");
            Map(m => m.Erzieher1Briefanrede).Name("Erzieher 1: Briefanrede");
            Map(m => m.Erzieher1Nachname).Name("Erzieher 1: Nachname");
            Map(m => m.Erzieher1Titel).Name("Erzieher 1: Titel");
            Map(m => m.Erzieher1Vorname).Name("Erzieher 1: Vorname");
            Map(m => m.Erzieher2Anrede).Name("Erzieher 2: Anrede");
            Map(m => m.Erzieher2Briefanrede).Name("Erzieher 2: Briefanrede");
            Map(m => m.Erzieher2Nachname).Name("Erzieher 2: Nachname");
            Map(m => m.Erzieher2Titel).Name("Erzieher 2: Titel");
            Map(m => m.Erzieher2Vorname).Name("Erzieher 2: Vorname");
            Map(m => m.ErzieherArtKlartext).Name("Erzieher: Art (Klartext)");
            Map(m => m.ErzieherEmail).Name("Erzieher: E-Mail");
            Map(m => m.ErzieherErhältAnschreiben).Name("Erzieher: Erhält Anschreiben");
            Map(m => m.ErzieherOrt).Name("Erzieher: Ort");
            Map(m => m.ErzieherOrtsteil).Name("Erzieher: Ortsteil");
            Map(m => m.ErzieherPostleitzahl).Name("Erzieher: Postleitzahl");
            Map(m => m.ErzieherStraße).Name("Erzieher: Straße");
            Map(m => m.ExterneIdNummer).Name("Externe ID-Nummer");
            Map(m => m.Fachklasse).Name("6 Fachklasse", "Fachklasse", "Fachklasse (Kürzel)");
            Map(m => m.FachklasseBezeichnung).Name("Fachklasse (Bezeichnung)");
            Map(m => m.FaxNr).Name("Allg. Adresse: Fax-Nr.", "Fax-Nr.");
            Map(m => m.Förderschwerpunkt1).Name("10 Foerderschwerp", "Förderschwerpunkt 1");
            Map(m => m.Förderschwerpunkt2).Name("Förderschwerpunkt 2", "63 Förderschwerpunkt 2");
            Map(m => m.Geburtsdatum).Name("Geburtsdatum", "16 Gebdat", "birthDate");
            Map(m => m.Geburtsland).Name("Geburtsland");
            Map(m => m.GeburtslandMutter).Name("Geburtsland Mutter", "54 Geb.Land (Mutter)");
            Map(m => m.GeburtslandVater).Name("Geburtsland Vater", "55 Geb.Land (Vater)");
            Map(m => m.Geburtsname).Name("Geburtsname");
            Map(m => m.Geburtsort).Name("Geburtsort");
            Map(m => m.Geschlecht).Name("gender", "Geschlecht", "17 Geschlecht");
            Map(m => m.Gliederung).Name("5 Gliederung", "Gliederung", "Schulgliederung");
            Map(m => m.GsEmpfehlung).Name("59 GS-Empfehlung");
            Map(m => m.Hausnummer).Name("Hausnummer");
            Map(m => m.HöchsterAllgAbschluss).Name("Höchster allg. Abschluss");
            Map(m => m.Internatsplatz).Name("69 Internatsplatz");
            Map(m => m.IdSchild).Name("Interne ID-Nummer");
            Map(m => m.JahrZuzug).Name("50 Jahr Zuzug");
            Map(m => m.JahrEinschulung).Name("51 Jahr Einschulung");
            Map(m => m.JahrSchulwechsel).Name("52 Jahr Schulwechsel");
            Map(m => m.Jahrgang).Name("Jahrgang");
            Map(m => m.JahrgangInterneBezeichnung).Name("Jahrgang (interne Bezeichnung)");
            Map(m => m.Jva).Name("13 Jva");
            Map(m => m.Klasse).Name("4 Klasse", "Klasse");
            Map(m => m.Klassenart).Name("7 Klassenart");
            Map(m => m.Klassenlehrer).Name("Klassenlehrer");
            Map(m => m.KlassenlehrerAmtsbezeichnung).Name("Klassenlehrer: Amtsbezeichnung");
            Map(m => m.KlassenlehrerAnrede).Name("Klassenlehrer: Anrede");
            Map(m => m.KlassenlehrerName).Name("Klassenlehrer: Name");
            Map(m => m.KlassenlehrerTitel).Name("Klassenlehrer: Titel");
            Map(m => m.KlassenlehrerVorname).Name("Klassenlehrer: Vorname");
            Map(m => m.Kreis).Name("Kreis");
            Map(m => m.Koopklasse).Name("70 Koopklasse");
            Map(m => m.LetzteSchuleName).Name("Letzte Schule: Name");
            Map(m => m.LetzteSchuleOrt).Name("Letzte Schule: Ort");
            Map(m => m.LetzteSchulePlz).Name("Letzte Schule: PLZ");
            Map(m => m.LetzterBerufsbezAbschlussKürzel).Name("Letzter berufsbez. Abschluss (Kürzel)");
            Map(m => m.LsSchulform).Name("26 LSSchulform", "Letzte Schule: Schulform");
            Map(m => m.Lsschulnummer).Name("27 Lsschulnummer", "Letzte Schule: Schulnr.");
            Map(m => m.LsGliederung).Name("28 LSGliederung");
            Map(m => m.LsFachklasse).Name("29 LSFachklasse");
            Map(m => m.Lsklassenart).Name("30 Lsklassenart");
            Map(m => m.Lsreformpdg).Name("31 Lsreformpdg");
            Map(m => m.LsSschulentl).Name("32 LSSschulentl", "Letzte Schule: Entlassdatum");
            Map(m => m.LsJahrgang).Name("33 LSJahrgang", "Letzte Schule: Entlassjahrgang");
            Map(m => m.LsQual).Name("34 LSQual", "Letzter allg. Abschluss (Kürzel)", "Letzte Schule: Abschluss");
            Map(m => m.Lsversetz).Name("35 Lsversetz", "Letzte Schule: Versetzungsvermerk");
            Map(m => m.MailPrivat).Name("address.email", "E-Mail (privat)", "Allg. Adresse: E-Mail");
            Map(m => m.MailSchulisch).Name("E-Mail schulisch)");
            Map(m => m.Massnahmetraeger).Name("60 Massnahmetraeger");
            Map(m => m.MigrationshintergrundVorhanden).Name("Migrationshintergrund vorhanden");
            Map(m => m.Nachname).Name("Nachname", " Familienname", "longName", "Allg. Adresse: Name1");
            Map(m => m.Orgform).Name("8 Orgform", "Orgform", "Organisationsform");
            Map(m => m.Ortsteil).Name("Ortsteil");
            Map(m => m.Ort).Name("Ortsname", "address.city", "15. Ort", "Allg. Adresse: Ort", "Ort");
            Map(m => m.Postleitzahl).Name("14 Plz", "Postleitzahl", "address.postCode", "Allg. Adresse: PLZ");
            Map(m => m.Produktname).Name("66 Produktname");
            Map(m => m.Produktversion).Name("67 Produktversion");
            Map(m => m.Reformpdg).Name("12 Reformpdg");
            Map(m => m.Religionsanmeldung).Name("20 Relianmeldung", "Datum Religionsanmeldung", "Religionsanmeldung");
            Map(m => m.Religionsabmeldung).Name("21 Reliabmeldung", "Religionsabmeldung", "Datum Religionsabmeldung");
            Map(m => m.KonfessionKlartext).Name("19 Religion");
            Map(m => m.Labk).Name("23 Labk");
            Map(m => m.Schwerstbehinderung).Name("Schwerstbehinderung", "11 Schwerstbehindert");
            Map(m => m.SchülerLeJahrDa).Name("Schüler le. Jahr da");
            Map(m => m.Schulwechselform).Name("48 Schulwechselform");
            Map(m => m.Schulbesuchsjahre).Name("Schulbesuchsjahre");
            Map(m => m.Schulform).Name("Schulform");
            Map(m => m.SchulformFürSimExport).Name("Schulform (f. SIM-Export)");
            Map(m => m.Schuljahr).Name("Schuljahr (z.B. 2008/2009)");
            Map(m => m.SchulNummer).Name("Schul-Nummer");
            Map(m => m.SchulpflichtErfüllt).Name("Schulpflicht erfüllt", "47 Schulpflichterf");
            Map(m => m.Schwerpunkt).Name("Schwerpunkt");
            Map(m => m.Sportbefreiung).Name("Sportbefreiung");
            Map(m => m.StaatsangehörigkeitKlartext).Name("Staatsangehörigkeit (Klartext)", "18 Staatsang");
            Map(m => m.StaatsangehörigkeitKlartextAdjektiv).Name("Staatsangehörigkeit (Klartext, Adjektiv)");
            Map(m => m.StaatsangehörigkeitSchlüssel).Name("Staatsangehörigkeit (Schlüssel)", "18 Staatsang");
            Map(m => m.Straße).Name("Straße", "Straßenname", "address.street", "Allg. Adresse: Straße");
            Map(m => m.Status).Name("2 Status", "Status");
            Map(m => m.Telefonnummer).Name("Telefon-Nr.", "Allg. Adresse: 1. Tel.-Nr.", "address.phone",
                "Telefon-Nummern: Telefon-Nummer");
            Map(m => m.Telefonnummer2).Name("Allg. Adresse: 2. Tel.-Nr.", "address.mobile");
            Map(m => m.TelefonNummernBemerkung).Name("Telefon-Nummern: Bemerkung");
            Map(m => m.Verkehrssprache).Name("57 Verkehrssprache", "Verkehrssprache in der Familie");
            Map(m => m.Vertragsbeginn).Name("Allg. Adresse: Vertragsbeginn");
            Map(m => m.Vertragsende).Name("Allg. Adresse: Vertragsende");
            Map(m => m.Versetzung).Name("49 Versetzung", "Versetzung");
            Map(m => m.VorjahrC05AktjahrC06).Name("Vorjahr C05 Aktjahr C06");
            Map(m => m.VoKlasse).Name("36 VOKlasse");
            Map(m => m.VoGliederung).Name("37 VOGliederung");
            Map(m => m.VoFachklasse).Name("38 VOFachklasse");
            Map(m => m.VoOrgform).Name("39 VOOrgform");
            Map(m => m.VoKlassenart).Name("40 VOKlassenart");
            Map(m => m.VoJahrgang).Name("41 VOJahrgang");
            Map(m => m.VoFoerderschwerp).Name("42 VOFoerderschwerp");
            Map(m => m.VoSchwerstbehindert).Name("43 VOSchwerstbehindert");
            Map(m => m.VoReformpdg).Name("44 VOReformpdg");
            Map(m => m.VoFoerderschwerpunkt2).Name("64 VOFörderschwerpunkt 2");
            Map(m => m.Volljährig).Name("Volljährig");
            Map(m => m.VoraussAbschlussdatum).Name("vorauss. Abschlussdatum");
            Map(m => m.Zeugnis).Name("46 Zeugnis");
            Map(m => m.Zugezogen).Name("53 zugezogen", "Zuzugsjahr");
        }
    }

    public void FotosAusAtlantisNachSuS()
    {

    }

    public List<dynamic> AdressenImportieren(Dateien dateien)
    {
        var records = new List<dynamic>();
        var adressen = dateien.FirstOrDefault(datei => datei.UnterordnerUndDateiname.Contains("Adressen"))!.ToList();

        foreach (var student in this)
        {
            var elternStattVundM = false;

            var adressenDiesesSchuelers = adressen.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Schüler: Vorname"].ToString() == student.Vorname &&
                       dict["Schüler: Nachname"].ToString() == student.Nachname &&
                       dict["Schüler: Geburtsdatum"].ToString() == student.Geburtsdatum.ToString();
            }).ToList();

            var vater = adressenDiesesSchuelers.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Adresse: Typ Adresse"].ToString() == "V";
            }).FirstOrDefault();

            var mutter = adressenDiesesSchuelers.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Adresse: Typ Adresse"].ToString() == "M";
            }).FirstOrDefault();

            if (vater != null && mutter != null)
            {
                var dictVater = (IDictionary<string, object>)vater;
                var dictMutter = (IDictionary<string, object>)mutter;

                // Wenn Mama und Papa dieselbe Adresse haben, ...
                if (dictVater!["Adresse: Straße"].ToString() == dictMutter!["Adresse: Straße"].ToString())
                {
                    // und der student nicht volljährig ist, ...
                    if (!student.IstVolljaehrig())
                    {
                        // ... wird ein Erzieher namens Eltern erzeugt, der als einziger angeschrieben wird.

                        dynamic record = new ExpandoObject();
                        record.Nachname = student.Nachname;
                        record.Vorname = student.Vorname;
                        record.Geburtsdatum = student.Geburtsdatum;
                        record.Erzieherart = "Eltern";
                        record.AnredeLEERZEICHEN1PUNKTPerson = "";
                        record.TitelLEERZEICHEN1PUNKTPerson = "";
                        record.NachnameLEERZEICHEN1PUNKTPerson = student.Nachname;
                        record.VornameLEERZEICHEN1PUNKTPerson =
                            dictMutter!["Adresse: Name 2"] + " und " + dictVater!["Adresse: Name 2"];
                        record.AnredeLEERZEICHEN2PUNKTPerson = "";
                        record.TitelLEERZEICHEN2PUNKTPerson = "";
                        record.NachnameLEERZEICHEN2PUNKTPerson = "";
                        record.VornameLEERZEICHEN2PUNKTPerson = "";
                        record.Straße = dictVater!["Adresse: Straße"].ToString();
                        record.PLZ = dictVater!["Adresse: Plz"].ToString();
                        record.Ort = dictVater!["Adresse: Ort"].ToString();
                        record.Ortsteil = "";
                        record.EMINUSMailLEERZEICHEN1PUNKTLEERZEICHENPerson = dictMutter!["Adresse: E-Mail"].ToString();
                        record.Anschreiben = "J";
                        record.EMINUSMailLEERZEICHEN2PUNKTLEERZEICHENPerson =
                            string.IsNullOrEmpty(dictVater!["Adresse: E-Mail"].ToString())
                                ? ""
                                : dictVater!["Adresse: E-Mail"].ToString();
                        records.Add(record);
                        elternStattVundM = true;
                    }
                }
            }

            foreach (var rec in adressenDiesesSchuelers)
            {
                var dict = (IDictionary<string, object>)rec;
                dynamic record = new ExpandoObject();
                record.Nachname = student.Nachname;
                record.Vorname = student.Vorname;
                record.Geburtsdatum = student.Geburtsdatum;
                record.Erzieherart = student.GetErzieherart(dict["Adresse: Typ Adresse"].ToString(),
                    dict["Adresse: Sorgeberechtigt (J/N)"].ToString(), dict["Schüler: Geschlecht"].ToString());
                record.AnredeLEERZEICHEN1PUNKTPerson = dict["Adresse: Anrede (Auflösung)"].ToString();
                record.TitelLEERZEICHEN1PUNKTPerson = "";
                record.NachnameLEERZEICHEN1PUNKTPerson = dict["Adresse: Name 1"].ToString();
                record.VornameLEERZEICHEN1PUNKTPerson = dict["Adresse: Name 2"].ToString();
                record.AnredeLEERZEICHEN2PUNKTPerson = "";
                record.TitelLEERZEICHEN2PUNKTPerson = "";
                record.NachnameLEERZEICHEN2PUNKTPerson = "";
                record.VornameLEERZEICHEN2PUNKTPerson = "";
                record.Straße = dict!["Adresse: Straße"].ToString();
                record.PLZ = dict!["Adresse: Plz"].ToString();
                record.Ort = dict!["Adresse: Ort"].ToString();
                record.Ortsteil = "";
                record.EMINUSMailLEERZEICHEN1PUNKTLEERZEICHENPerson = dict!["Adresse: E-Mail"].ToString();
                record.Anschreiben = Anschreiben(record.Erzieherart, student.IstVolljaehrig(), elternStattVundM);
                record.EMINUSMailLEERZEICHEN2PUNKTLEERZEICHENPerson = "";
                if (!(!student.IstVolljaehrig() && record.Erzieherart == ""))
                {
                    records.Add(record);
                }
            }
        }

        return records;
    }

    private static dynamic Anschreiben(dynamic erzieherart, bool volljaehrig, bool elternStattVundM)
    {
        // Minderjährige werden nicht informiert
        if (!volljaehrig && erzieherart == "")
        {
            return "N";
        }

        // Volljährige immer
        if (volljaehrig && erzieherart.StartsWith("Schüler"))
        {
            return "J";
        }

        if (erzieherart == "Eltern")
        {
            return "J";
        }

        // Vater oder Mutter Minderjähriger, wenn es Eltern gibt
        if (!volljaehrig && erzieherart != "" && elternStattVundM)
        {
            return "N";
        }

        // Vater oder Mutter Minderjähriger, ohne Eltern
        if (!volljaehrig && erzieherart != "" && !elternStattVundM)
        {
            return "J";
        }

        return "N";
    }

    public void FotosNachSchildSchreiben(List<Klasse> iKlassen, IConfiguration configuration)
    {
        DateTime[] releases =
        {
            new DateTime(2017, 2, 11),
            new DateTime(2017, 5, 11),
            new DateTime(2017, 11, 14),
            new DateTime(2017, 10, 14),
            new DateTime(2018, 02, 02),
            new DateTime(2018, 06, 10),
            new DateTime(2025, 01, 18)
        };

        SchildVersionExpected = releases;

        Global.Konfig("MariaDb", configuration);
        Global.Konfig("MariaIp", configuration);
        Global.Konfig("MariaPort", configuration);
        Global.Konfig("MariaUser", configuration);
        Global.Konfig("MariaPw", configuration);

        var dataAccess = new DataAccess();
        var schildVersionActual = dataAccess.GetSchildVersion();

        var students = dataAccess.GetStudents(iKlassen[0]);

        if (SchildVersionExpected.All(x => x.Date < schildVersionActual.Date))
        {
            try
            {
                System.Diagnostics.Process.Start("https://github.com/BerufskollegBorken/schuelerfoto");
            }
            catch (Exception)
            {
            }

            Console.WriteLine(
                "Diese Version von SchuelerFoto-NRW ist aktuell nur freigegeben für die SchILD-Version vom " +
                SchildVersionExpected.Max().Date.ToShortDateString() +
                ".\n\rIhre SchILD-Version: " + schildVersionActual.Date.ToShortDateString() +
                ". Schauen Sie hier: https://github.com/BerufskollegBorken/schuelerfoto");
        }

        int nr = 1;
        var pics = new List<string>();

        do
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("  Ziehen Sie genau " + students.Count() +
                              " Schülerfotos in folgender Reihenfolge aus dem Explorer auf die App:");
            Console.ResetColor();
            nr = 0;
            foreach (var student in students)
            {
                nr++;
                Console.WriteLine(nr.ToString().PadLeft(4) + ". " + student.Nachname + "," + student.Vorname);
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("  Danach ENTER:");
            Console.ResetColor();
            var picsString = Console.ReadLine();
            pics = new List<string>();

            // Regex für Pfade in einfachen oder doppelten Anführungszeichen
            string pattern = @"['""]([^'""]+)['""]"; // Alles zwischen einfachen oder doppelten Quotes

            // Regex.Match für mehrere Gruppen
            MatchCollection matches = Regex.Matches(picsString, pattern);

            foreach (Match match in matches)
            {
                if (match.Groups.Count > 1) // Überprüfen, ob ein Pfad gefunden wurde
                {
                    pics.Add(match.Groups[1].Value);
                }
            }

            if (pics.Count() == nr)
            {
                break;
            }

            Console.ForegroundColor = ConsoleColor.Red;

            Console.WriteLine("   Es müssen exakt " + students.Count() + " Bilder hochgeladen werden. Es wurden " +
                              pics.Count() + " Bilder hochgeladen.");
            Console.WriteLine(
                "   Die Anzahl stimmt nicht überein. Es dürfen nur jpg oder jpeg-Dateien hochgeladen werden.");
            Console.WriteLine("");
            Console.ResetColor();
        } while (true);

        Console.WriteLine("  Es wurden " + students.Count() + " hochgeladen. Gut so.");

        for (int i = 0; i < pics.Count; i++)
        {
            Console.WriteLine("");
            Console.Write(" " + Path.GetFileName(pics[i]) + " ... ");

            // Beispiel: Bild laden und bearbeiten
            using (var image = Image.Load(pics[i]))
            using (var ms = new MemoryStream()) // MemoryStream für die Umwandlung
            {
                image.Mutate(x => x.Resize(160, 160));
                Console.Write("160x160 ...");
                //image.Save("path/to/output.jpg");
                image.SaveAsJpeg(ms); // Bild als JPEG in den Stream speichern
                byte[] imageBytes = ms.ToArray(); // Stream in ein Byte-Array umwandeln
                students[i].Foto = Convert.ToBase64String(imageBytes); // Byte-Array in Base64 umwandeln

                dataAccess.DeleteImage(students[i]);
                dataAccess.InsertImage(students[i]);
            }
        }
    }

    public List<dynamic> AdressenBetriebeImportieren(Dateien dateien)
    {
        var records = new List<dynamic>();
        List<dynamic> adressen =
            dateien.FirstOrDefault(datei => datei.UnterordnerUndDateiname.Contains("Adressen"))!.ToList();
        List<dynamic> datums =
            dateien.FirstOrDefault(datei => datei.UnterordnerUndDateiname.Contains("DatumsAusAtlantis"))!.ToList();

        foreach (var student in this)
        {
            var adressenDiesesSchuelers = adressen.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Schüler: Vorname"].ToString() == student.Vorname &&
                       dict["Schüler: Nachname"].ToString() == student.Nachname &&
                       dict["Schüler: Geburtsdatum"].ToString() == student.Geburtsdatum.ToString();
            }).ToList();

            var datumsDiesesSchuelers = datums.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Vorname"].ToString() == student.Vorname &&
                       dict["Nachname"].ToString() == student.Nachname &&
                       DateTime.Parse(dict["Geburtsdatum"].ToString()).ToString("dd.MM.yyyy") ==
                       student.Geburtsdatum.ToString();
            }).LastOrDefault();


            var dictD = (IDictionary<string, object>)datumsDiesesSchuelers;

            foreach (var rec in adressenDiesesSchuelers)
            {
                var dictA = (IDictionary<string, object>)rec;

                if (!string.IsNullOrEmpty(dictA["Betrieb: Name 1"].ToString()))
                {
                    dynamic record = new ExpandoObject();
                    record.Nachname =
                        student.Nachname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                    record.Vorname =
                        student.Vorname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                    record.Geburtsdatum =
                        student.Geburtsdatum; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                    record.Adressart = "Betrieb";
                    record.Name1 = dictA["Betrieb: Name 1"].ToString();
                    record.Name2 = dictA["Betrieb: Name 2"].ToString();
                    record.Straße = dictA["Betrieb: Straße"].ToString();
                    record.PLZ = dictA["Betrieb: Plz"].ToString();
                    record.Ort = dictA["Betrieb: Ort"].ToString();
                    record.EINSPUNKTLEERZEICHENTelPUNKTMINUSNrPUNKT = dictA["Betrieb: Telefon 1"].ToString();
                    record.ZWEIPUNKTLEERZEICHENTelPUNKTMINUSNrPUNKT = dictA["Betrieb: Telefon 2"].ToString();
                    record.EMINUSMail = dictA["Betrieb: E-Mail-Adresse"].ToString();
                    record.BetreuerLEERZEICHENNachname = dictA["Betrieb: Ansprechpartner"].ToString();
                    record.BetreuerLEERZEICHENVorname = "";
                    record.BetreuerLEERZEICHENAnrede = "";
                    record.BetreuerLEERZEICHENTelPUNKTMINUSNrPUNKT = "";
                    record.BetreuerLEERZEICHENEMINUSMail = "";
                    record.BetreuerLEERZEICHENAbteilung = "";
                    record.Vertragsbeginn = dictD != null && !string.IsNullOrEmpty(dictD["Ausbildungsbeginn"]?.ToString())
                        ? DateTime.Parse(dictD["Ausbildungsbeginn"]?.ToString() ?? string.Empty).ToString("dd.MM.yyyy")
                        : "";
                    record.Vertragsende = dictD != null && !string.IsNullOrEmpty(dictD["Ausbildungsende"]?.ToString())
                        ? DateTime.Parse(dictD["Ausbildungsende"]?.ToString() ?? string.Empty).ToString("dd.MM.yyyy")
                        : "";

                    var gibtEsSchon = records.Where(recB =>
                    {
                        var dictB = (IDictionary<string, object>)recB;
                        return dictB["Nachname"].ToString() == dictA["Schüler: Nachname"].ToString() &&
                               dictB["Vorname"].ToString() == dictA["Schüler: Vorname"].ToString() &&
                               dictB["Name1"].ToString() == dictA["Betrieb: Name 1"].ToString();
                    }).Any();

                    if (!gibtEsSchon)
                    {
                        records.Add(record);
                    }
                }
            }
        }

        return records;
    }

    public Klassen GetKlassen(Klassen klassen)
    {
        foreach (var student in this)
        {
            if (!klassen.Any(x => x.Name == student.KlasseWebuntis))
            {
                klassen.Add(new Klasse { Name = student.KlasseWebuntis });
            }

            if (!klassen.Any(x => x.Name == student.Klasse))
            {
                klassen.Add(new Klasse { Name = student.Klasse });
            }
        }

        Klassen kla = new Klassen();
        kla.AddRange(klassen.Where(x => !string.IsNullOrEmpty(x.Name)));
        return kla;
    }

    public Students GetInteressierende(string klasse)
    {
        var students = new Students();
        foreach (var stu in this)
        {
            if (klasse.Contains("alle") || klasse == stu.Klasse || klasse == stu.KlasseWebuntis)
            {
                students.Add(stu);
            }
        }

        return students;
    }

    public bool IsAllesOk()
    {
        var stati = this.Select(x => x.Status).Distinct().ToList();

        if (Count > 0 && stati.Count > 1)
        {
            return true;
        }

        return false;
    }

    public Students DoppelteFiltern()
    {
        var stdents = new Students();
        
        // Wenn SuS wiederkommen usw., dann bekommen sie eine neue, höhere ID. Also gewinnt die höchste ID.
        foreach (var student in this.OrderByDescending(x=>x.IdSchild))
        {
            if (!stdents.Where(x=>x.Nachname == student.Nachname && x.Vorname == student.Vorname && x.Geburtsdatum == student.Geburtsdatum).Any())
            {
                stdents.Add(student);
            }
            else
            {
                string doppelt;
            }
        }
        return stdents;
    }
}