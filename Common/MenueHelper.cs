using Common;
using Microsoft.Extensions.Configuration;

public static class MenueHelper
{
    public static Menue Einlesen(Dateien quelldateien, Klassen Klassen, Lehrers lehrers, IConfiguration configuration, Anrechnungen anrechnungen,
        Raums raums)
    {
        //  Die Schüler werden aus dem SchildSchuelerExport eingelesen
    
        var students = new Students(quelldateien);

        var zieldatei = new Datei();
        return new Menue(
            quelldateien,
            Klassen,
            lehrers,
            students,
            [
                new Menüeintrag(
                    "Klassen neu anlegen",
                    anrechnungen,
                    quelldateien.Notwendige(["klassen", "GPU003"]),
                    students,
                    Klassen,
                    [
                        "Es wird angenommen, dass die Klassen des kommenden Schuljahres zuerst in Untis angelegt werden.",
                        "Alle neuen Klassen in Untis können vor dem Schuljahreswechsel nach SchILD übergeben werden.",
                        "Dazu die GPU003.TXT aus Untis exportieren, um neue Klassenbezeichnungen zu identifizieren.",
                        "Eigenschaften der neuen Klassen werden aus SchILD-Vorjahresklassen entnommen.",
                        "Bei vorhandenen Klassen werden abweichende Eigenschaften (z.B. Klassenleitung) angepasst.",
                    ],
                    m =>
                    {
                        zieldatei = m.KlassenErstellen(Path.Combine(Global.PfadSchilddateien ?? "", @"Klassen.dat"));
                        if (!zieldatei.Any()) return;
                        zieldatei = zieldatei.VergleichenUndFiltern(quelldateien, ["InternBez"], ["SonstigeBez", "Folgeklasse"]);
                        zieldatei?.Erstellen("|", '\0', true, false);
                    }
                ),
                new Menüeintrag(
                    "Webuntis-Export-Datei erstellen (Schüler*innen plus Fotos)",
                    anrechnungen,
                    quelldateien.Notwendige([
                        "student_", "schuelerzusatzdaten", "schuelererzieher", "schuelerbasisdaten", "schuelerAdressen"
                    ]),
                    students,
                    Klassen,
                    [
                        "Importdateien für Webuntis werden erstellt:",
                        "1. Die Datei Student_***.csv muss als admin aus Webuntis nach " + Global.PfadExportdateien + " heruntergeladen werden: Stammdaten->Schüler->Export als csv",
                        "2. Aus SchILD werden SchuelerBasisdaten, Schuelerzusatzdaten, SchuelerErzieher und SchuelerAdressen benötigt" +
                        "3. Alle Schüler*innen, die nicht mehr aktiv (Status 2) oder Gast (Status 6) sind und kein Austrittsdatum haben, bekommen den heutigen Tag als Austrittsdatum gesetzt",
                        "4. Die Fotos der neuen Schülerinnen und Schüler werden in eine Zip-Datei gepackt"
                    ],
                    m =>
                    {
                        zieldatei = m.WebuntisOderNetmanCsv(Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") +  @"-ImportNachWebuntis.csv"));
                        zieldatei?.Erstellen(";", '\'', false, false);                        
                    }
                ),
                new Menüeintrag(
                    "Netman und Littera: Exportdateien erstellen",
                    anrechnungen,
                    quelldateien.Notwendige([
                        "student_", "schuelerzusatzdaten", "schuelererzieher", "schuelerbasisdaten", "klassen", "lehrkraefte"
                    ]),
                    students,
                    Klassen,
                    [
                        "Importdateien für Netman und Littera werden erstellt:",
                        "1. Die Datei Student_***.csv muss als admin aus Webuntis nach " + Global.PfadExportdateien + " heruntergeladen werden: Stammdaten->Schüler-> Export als csv.",
                        "2. Nach dem Einlesen werden alle Schülerinnen und Schüler angezeigt, die in Untis zu löschen sind. Es wird bei diesen Schülerinnen und Schülern ein Austrittsdatum gesetzt.",
                        "3. Die Importdatei für Webuntis wird geschrieben.",
                        "4. Die Fotos der neuen Schülerinnen und Schüler werden in eine Zip-Datei gepackt."
                    ],
                    m =>
                    {                      
                        zieldatei = m.WebuntisOderNetmanCsv(Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.AddHours(1).ToString("yyyyMMdd-HHmm") + @"-ImportNachNetman.csv"));
                        zieldatei?.Erstellen(",", '\'', false, false);
                        zieldatei?.Zippen(zieldatei?.GetAbsoluterPfad());
                        zieldatei?.Mailen(Path.GetFileName(zieldatei.AbsoluterPfad) ?? "", "Verwaltung", Path.GetFileName(zieldatei.AbsoluterPfad) ?? "", Global.NetmanMailReceiver ?? "");

                        zieldatei = m.WebuntisOderNetmanCsv(Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.AddHours(1).ToString("yyyyMMdd-HHmm") + @"-ImportNachLittera.csv"));
                        zieldatei?.Erstellen(";", '\'', false, false);
                        zieldatei?.Zippen(zieldatei?.GetAbsoluterPfad());
                        zieldatei?.Mailen(Path.GetFileName(zieldatei.AbsoluterPfad ?? ""), "Verwaltung", Path.GetFileName(zieldatei.AbsoluterPfad ?? ""), Global.NetmanMailReceiver  ?? "");
                    }
                ),
                new Menüeintrag(
                    "Outlook: Termine für Wiki aufbereiten",
                    anrechnungen,
                    quelldateien.Notwendige([
                        "termine_fhr", "termine_verwaltung", "termine_berufliches_gymnasium", "termine_kollegium"
                    ]),
                    students,
                    Klassen,
                    [
                        "Die Kalender müssen mit Copy&Paste aus Outlook in die CSV-Dateien im Download-Ordner kopiert werden.",
                        "Falls der Inhalt im Body (Spalte Nachricht) mehrzeilig ist, wird nur die erste Zeile berücksichtigt.",
                        "Es werden nur Termine berücksichtigt, die mindestens eine Kategorie haben. Kategorien werden zu Links in Wiki.",
                        "Termine aus vergangenen Schuljahren werden nicht mit übertragen.",                        
                        "Die Kalender im Wiki zuerst leeren. Anschließend die neuen CSV als Global importieren."
                    ],
                    m =>
                    {
                        foreach (var kalender in new List<string>(){"termine_berufliches_gymnasium", "termine_kollegium", "termine_verwaltung", "termine_fhr" })
                        {
                            zieldatei = m.Kalender2Wiki(kalender, Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachWiki-" + kalender));
                            zieldatei.Erstellen(",", '\"', false, true);    
                        }
                    }
                ),
                new Menüeintrag(
                    "Sprechtag vorbereiten: Alle Lehrkräfte mit eigenem Unterricht werden auf die Wiki-Seite geschrieben",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,                    
                    [
                        "Die Wiki-Datei sprechtag.txt wird angepasst. Die Wunschräume werden in den Untis-Stammdaten beim Lehrer eingetragen. Dazu die Fenstergruppe Sprechtag in Untis öffnen. Bei Abwesenheiten die Räume für kommendes Jahr stehen lassen, wenn im Betreff 'außer Haus' steht, dann wird der Raum nicht angezeigt. Fußnoten werden als Text2 in den Untis-Stammdaten eingetragen. Beispiel für eine Fußnote: 'außer Haus; bitte Termin vereinbaren;'",
                        "Lehrkräfte ohne Raum werden in der Liste ignoriert. Lehrkräfte ohne eigenen Unterricht bleiben unberücksichtigt"
                    ],
                    m =>
                    {
                        m.Sprechtag(lehrers, raums, configuration,
                            "Zum jährlichen Sprechtag laden wir sehr herzlich am Mittwoch nach der Zeugnisausgabe in der Zeit von 13:30 bis 17:30 Uhr ein. Der Unterricht endet nach der 5. Stunde um 12:00 Uhr.");
                    }
                ),
                new Menüeintrag(
                    "BKB.wiki: Dateien erstellen",
                    anrechnungen,
                    quelldateien.Notwendige([
                        "schuelerZusatzdaten", "absenceperstudent"
                    ]),
                    students,
                    Klassen,
                    [
                        "Das Organigramm wird aus Untisanrechnungen gebildet. Beispiele: {...} --> KATEGORIE; [...] --> HINWEIS, Text ohne Klammern wird zur ROLLE; A14, A15, A16 ohne Klammern --> AMT; Untis-Beschreibung --> AUFGABE. Im Organigramm wird nach Kategorie, Aufgabe oder Beschreibung gruppiert.",
                        "Untisanrechnungen: 1.Struct Schema Editor > Untisanrechnungen > Löschen/Leeren > 'untisanrechnungen' eingeben, dann Leeren",
                        "Untisanrechnungen: 2.Struct Schema Editor > Untisanrechnungen > Importieren/Exportieren > Importieren von Rohdaten > Global > Durchsuchen"
                    ],
                    m =>
                    {
                        zieldatei = m.GetGruppen(Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-gruppen.csv"), anrechnungen, lehrers);
                        zieldatei.Erstellen(",", '\"', false, true);
                        zieldatei = anrechnungen.Anlegen(Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-untisanrechnungen.csv") ,[500, 510, 530, 590, 900], [500, 510, 530, 590], ["PLA", "BM"]);
                        zieldatei.Erstellen(",", '\"', false, true);

                        zieldatei = m.GetLehrer(Path.Combine(Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-lul-utf8OhneBom-einmalig-vor-SJ-Beginn.csv")));
                        zieldatei.Erstellen(",", '\'', false, false);

                        zieldatei = m.Praktikanten(
                            [
                                "BW,1", "BT,1", "BS,1", "BS,2", "HBG,1", "HBT,1", "HBW,1", "GG,1", "GT,1", "GW,1",
                                "IFK,1"
                            ],
                            Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-praktikanten-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"));
                        zieldatei?.Erstellen(",", '\'', false, false);

                        zieldatei = m.KlassenAnlegen(Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-klassen-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"));
                        zieldatei?.Erstellen(",", '\'', false, false);

                        m.Schulpflichtüberwachung();

                        zieldatei = m.GetFaecher(Path.Combine(Global.PfadSchilddateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-faecher.csv"));
                        zieldatei?.Erstellen(",", '\'', false, false);
                    }
                ),
                new Menüeintrag(
                    "Leistungsdaten (Unterrichte, Fehlzeiten & Noten) für die Zeugnisschreibung oder UVD nach Schild importieren",
                    anrechnungen,
                    quelldateien.Notwendige(["schuelerlernabschnitt", "schuelerleistungsdaten", "schuelerbasis"]),
                    students,
                    Klassen,
                    [
                        "Die Leistungsdaten werden nach Schild importiert. Dazu gehören alle Unterrichte (mit Noten), sowie Kurse und Fächer. Lernabschnittsdaten mit Fehlzeiten gehören ebenfalls dazu. Es empfiehlt sich die Lernabschnitte zuerst in SchILD anzulegen und zu exportieren. Schule.exe ergänzt dann die Fehlzeiten passend.",
                        "Falls mehrere Kollegen dasselbe Fach zeitgleich unterrichten, dann muss ein Zähler an das Fach angehangen werden. Bsp: L1 -> M und L2 -> M1.",
                        "Damit M1 in den Leistungsdaten erscheint, aber nicht auf dem Zeugnis gedruckt wird, muss die Eigenschaft 'Nicht auf Zeugnis drucken' in SchILD gesetzt werden.",
                        "..."
                    ],
                    m =>
                    {
                        m.FilterInteressierendeStudentsUndKlassen();

                        zieldatei = m.Lernabschnittsdaten(Path.Combine(Global.PfadSchilddateien ?? "",
                            "SchuelerLernabschnittsdaten.dat"));
                        zieldatei = zieldatei.VergleichenUndFiltern(quelldateien,
                            ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"], []);
                        zieldatei?.Erstellen("|", '\0', true, false);

                        zieldatei = m.Leistungsdaten(Path.Combine(Global.PfadSchilddateien ?? "",
                            "SchuelerLeistungsdaten.dat"));
                        zieldatei = zieldatei.VergleichenUndFiltern(quelldateien,
                            ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt", "Fach"], ["Jahrgang"]);
                        zieldatei?.Erstellen("|", '\0', true, false);

                        zieldatei = m.Faecher(Path.Combine(Global.PfadSchilddateien ?? "", "Faecher.dat"));
                        zieldatei?.Erstellen("|", '\0', true, false);

                        zieldatei = m.Kurse(Path.Combine(Global.PfadSchilddateien ?? "", "Kurse.dat"));
                        zieldatei = zieldatei.VergleichenUndFiltern(quelldateien, ["KursBez"],
                            ["Klasse", "Schulnr", "WochenstdPUNKTLEERZEICHENKL"]);
                        zieldatei?.Erstellen("|", '\0', true, false);

                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("  Schritt #1: Import nach SchILD durchführen.");
                        Console.WriteLine("  Schritt #2: Notenkontrollliste nach Wiki hochladen.");
                        Console.WriteLine("  Schritt #3: Lehrkräfte benachrichtigen.");
                        Console.ResetColor();
                    }
                ),
                new Menüeintrag(
                    "Lehrkräfte an die Eintragung der Zeugnisnoten erinnern",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [
                        "Zuerst müssen die Leistungsdaten nach SchILD importiert werden.",
                        "Aus dem anschließenden Datenexport aus SchILD werden hier alle fehlenden Zeugnisnoten aufgelistet.",
                        "Alle betreffenden Mail-Adressen der Lehrkräfte können mit Copy & Paste nach Teams kopiert werden."
                    ],
                    m => m.LuLAnEintragungDerZeugnisnotenErinnern(lehrers)
                ),
                new Menüeintrag(
                    "Mahnungen importieren",
                    anrechnungen,
                    quelldateien.Notwendige(["marksperlesson"]),
                    students,
                    Klassen,
                    [
                        "Importdateien Mahnungen importieren."
                    ],
                    m =>
                    {
                        //dateien.Funktion(dateien.Benötigte(["MarksPerLessons"]));
                    }
                ),
                new Menüeintrag(
                    "Teams-Chat mit Lehrkräften beginnen",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [
                        "Lehrkräfte können über Teams angeschrieben werden."
                    ],
                    m =>
                    {
                        m.FilterInteressierendeStudentsUndKlassen();

                        var datei = m.GetGruppen("", anrechnungen, lehrers ?? []);

                        int index = -1;
                        while (true)
                        {
                            for (int i = 0; i < datei.Gruppen.Count; i++)
                            {
                                Console.WriteLine((i + 1).ToString().PadLeft(3) + ". " + datei.Gruppen[i].Name);
                            }

                            Console.Write("Bitte eine Zahl auswählen: ");
                            string? input = Console.ReadLine();

                            if (int.TryParse(input, out index) && index > 0 && index <= datei.Gruppen.Count)
                            {
                                break; // Gültige Eingabe, Schleife beenden
                            }
                            else
                            {
                                Console.WriteLine("Ungültige Eingabe! Bitte eine gültige Zahl eingeben.");
                            }
                        }

                        Console.WriteLine("Sie haben gewählt: " + datei.Gruppen[index - 1].Name);


                        m.ChatErzeugen(datei.Gruppen[index].Lehrers);
                    }
                ),
                new Menüeintrag(
                    "Kursbelegungen vorbereiten",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [
                        " 1. Alle Gymklassen der Jahrgangsstufen 12 und 13 aus SchILD exportieren.",
                        " 2. Alle Dateien in die Exceldatei namens Kursbelegung.xlsx importieren. Siehe LiesMich.",
                        " 3. Nachdem die Datei Kursbelegung.xlsx gefüllt wurde, die Datei nach SchILD importieren."
                    ],
                    m =>
                    {
                        //dateien.Kursbelegung(dateien.Benötigte([]));
                    }
                ),
                new Menüeintrag(
                    "Fotos der Schüler*innen klassenweise nach SchILD hochladen",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [
                        "Vorarbeiten: Fotos aller Schüler (z.B. mit dem Handy) erstellen. Dabei die Reihenfolge der Schüler laut SchILD exakt einhalten. ",
                        "   Hinweis #1: Wenn ein Schüler fehlt, dann die weiße Wand fotografieren, damit Anzahl und Reihenfolge stimmen.",
                        "   Hinweis #2: Wenn ein Foto nicht gelungen ist, dann löschen und neu erstellen.",
                        "   Hinweis #3: Wenn mehr als eine Klasse ausgewählt wird, wird nur die erste Klasse berücksichtigt",
                        "Verarbeitung: Klasse auswählen und Bilder per Drag & Drop in auf die App ziehen."
                    ],
                    m =>
                    {
                        //m.Students.FotosAusAtlantisNachSuS();
                        m.Students.FotosNachSchildSchreiben(m.Klassen, configuration);
                        m.Students.Fotos();
                    }
                ),
                new Menüeintrag(
                    "SchuelerTeilleistungen nach Schild importieren",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [
                        "Die Teilleistungen werden von Webuntis nach SchILD übertragen",
                        "...",
                        "..."
                    ],
                    m =>
                    {
                        zieldatei = m.Teilleistungen(@"ImportNachSchild\SchuelerTeilleistungen.dat");
                        zieldatei.Erstellen("|", '\0', true, false);
                    }
                ),
                new Menüeintrag(
                    "Von PDF-Dateien auf dem Desktop verschlüsselte Kopien erstellen.",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    ["Die Funktion kann genutzt werden, um PDF-Dateien mit Kennwort in Office365 zu versenden."],
                    _ =>
                    {
                        var pdfDateien = new PdfDateien();
                        pdfDateien.KennwortSetzen();
                    }
                ),
                new Menüeintrag(
                    "Schnellmeldung September",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [
                        "Dokumentation siehe Schips.webuntis2schildGui.nrw.de",
                        "Realtionen gemäß §93 SchulG"
                    ],
                    _ => { new Relationsgruppen(Klassen, students); }
                ),
                new Menüeintrag(
                    "Altersermäßigung berechnen",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [],
                    _ => lehrers.CheckAltersermäßigung(anrechnungen)                    
                ),
                new Menüeintrag(
                    "Lernabschnittsdaten und Leistungsdaten alter Abschnitte von Atlantis nach SchILD importieren",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [
                        "Die Lernabschnittsdaten (ohne Fehlzeiten und ohne Zeugnisdatum) und die Leistungsdaten alter Abschnitte werden für SchILD bereitgestellt."
                    ],
                    m =>
                    {
                        zieldatei = m.LernabschnittsdatenAlt(@"DatenaustauschSchild/SchuelerLernabschnittsdaten.dat");
                        zieldatei.Erstellen("|", '\0', true, false);

                        zieldatei = m.LeistungsdatenAlt(@"DatenaustauschSchild/SchuelerLeistungsdaten.dat");
                        zieldatei?.Erstellen("|", '\0', true, false);
                    }
                ),
                new Menüeintrag(
                    "PDF-Zeugnisse einsortieren",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [
                        "PDF-Zeugnisse und andere PDF-Dateien werden eingelesen.",
                        "Jede Seite wird nach Schülername durchsucht."
                    ],
                    _ =>
                    {
                        var suchmuster = new List<string>()
                        {
                            @"\b\w*Abgangszeugnis\w*\b",
                            @"\b\w*Abschlusszeugnis\w*\b",
                            @"\b\w*Jahreszeugnis\w*\b",
                        };

                        Global.Konfig("InputFolder", configuration);
                        Global.Konfig("OutputFolder", configuration);
                        new PdfDateien(students, suchmuster);
                    }
                ),
                new Menüeintrag(
                    "Zusatzdaten",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [],
                    m =>
                    {
                        zieldatei = m.Zusatzdaten(@"ImportNachSchild/SchuelerZusatzdaten.dat");
                        zieldatei.Erstellen("|", '\0', true, false);
                    }
                    ),
                new Menüeintrag(
                    "Basisdaten",
                    anrechnungen,
                    quelldateien,
                    students,
                    Klassen,
                    [],
                    m =>
                    {
                        zieldatei = m.Basisdaten("ImportNachSchild/SchuelerBasisdaten.dat");
                        zieldatei.Erstellen("|", '\0', true, false);
                    }
                )
            ]);
    }
}
    