using Common;
using Microsoft.Extensions.Configuration;
using PdfSharp.Pdf;

public static class MenueHelper
{
    public static Menue Einlesen(Dateien quelldateien, Klassen Klassen, Lehrers lehrers, IConfiguration configuration, Anrechnungen anrechnungen, Raums raums)
    {
        try
        {
            var students = new Students(quelldateien.Notwendige(["SchildSchuelerExport"]));
        
            if (students.Count == 0)
            {
                throw new Exception("Keine Schülerdaten gefunden.");
            }

            if(lehrers.Count == 0)
            {
                lehrers = new Lehrers(quelldateien.Notwendige(["lehrkraefte"]));
                    
                if (lehrers.Count == 0)
                {
                    throw new Exception("Keine Lehrerdaten gefunden.");
                }
            }

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
                            m.Zieldatei = m.KlassenErstellen(Path.Combine(Global.PfadSchilddateien ?? "", @"Klassen.dat"));
                            if (!m.Zieldatei.Any()) return;
                            m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, ["InternBez"], ["SonstigeBez", "Folgeklasse"]);
                            m.Zieldatei?.Erstellen("|", '\0', true, false);
                        }
                    ),
                    new Menüeintrag(
                        "Webuntis-Datei (Schüler*innen plus Fotos) erstellen",
                        anrechnungen,
                        quelldateien.Notwendige([
                            "student_", "schuelerzusatzdaten", "schuelererzieher", "schuelerbasisdaten", "schuelerAdressen", "lehrkraefte", "klassen" 
                        ]),
                        students,
                        Klassen,
                        [
                            "Importdateien für Webuntis werden erstellt:",
                            "1. Die erzeugte Datei Student_***.csv im Ordner " + Global.PfadExportdateien + " muss als admin nach Webuntis importiert werden",
                            "2. Stammdaten->Schüler*innen->Import",
                            "3. Datei auswählen, UTF8",
                            "4. Profil: Schuelerimport, dann Vorschau"
                        ],
                        m =>
                        {
                            m.Zieldatei = m.WebuntisOderNetmanCsv(Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") +  @"-ImportNachWebuntis.csv"));
                            m.Zieldatei?.Erstellen(";", '\'', false, false);                     
                        }
                    ),
                    new Menüeintrag(
                        "Netman und Littera: Exportdateien erstellen",
                        anrechnungen,
                        quelldateien.Notwendige([
                            "student_", "schuelerzusatzdaten", "schuelererzieher", "schuelerbasisdaten", "schuelerAdressen", "lehrkraefte", "klassen"
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
                            m.Zieldatei = m.WebuntisOderNetmanCsv(Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.AddHours(1).ToString("yyyyMMdd-HHmm") + @"-ImportNachNetman.csv"));
                            m.Zieldatei?.Erstellen(",", '\'', false, false);
                            m.Zieldatei?.Zippen(m.Zieldatei?.GetAbsoluterPfad(), configuration);
                            m.Zieldatei?.Mailen(Path.GetFileName(m.Zieldatei.AbsoluterPfad) ?? "", "Verwaltung", Path.GetFileName(m.Zieldatei.AbsoluterPfad) ?? "", Global.NetmanMailReceiver ?? "", configuration);

                            m.Zieldatei = m.WebuntisOderNetmanCsv(Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.AddHours(1).ToString("yyyyMMdd-HHmm") + @"-ImportNachLittera.csv"));
                            m.Zieldatei?.Erstellen(";", '\'', false, false);
                            m.Zieldatei?.Zippen(m.Zieldatei?.GetAbsoluterPfad(), configuration);
                            m.Zieldatei?.Mailen(Path.GetFileName(m.Zieldatei.AbsoluterPfad ?? ""), "Verwaltung", Path.GetFileName(m.Zieldatei.AbsoluterPfad ?? ""), Global.NetmanMailReceiver  ?? "", configuration);
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
                                m.Zieldatei = m.Kalender2Wiki(kalender, Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachWiki-" + kalender));
                                m.Zieldatei.Erstellen(",", '\"', false, true);    
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
                            "schuelerzusatzdaten", "absenceperstudent", "exportlesson"
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
                            m.Zieldatei = m.GetGruppen(Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-gruppen.csv"), anrechnungen, lehrers);
                            m.Zieldatei.Erstellen(",", '\"', false, true);
                            m.Zieldatei = anrechnungen.Anlegen(Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-untisanrechnungen.csv") ,[500, 510, 530, 590, 900], [500, 510, 530, 590], ["PLA", "BM"]);
                            m.Zieldatei.Erstellen(",", '\"', false, true);

                            m.Zieldatei = m.GetLehrer(Path.Combine(Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-lul-utf8OhneBom-einmalig-vor-SJ-Beginn.csv")));
                            m.Zieldatei.Erstellen(",", '\'', false, false);

                            m.Zieldatei = m.Praktikanten(
                                [
                                    "BW,1", "BT,1", "BS,1", "BS,2", "HBG,1", "HBT,1", "HBW,1", "GG,1", "GT,1", "GW,1",
                                    "IFK,1"
                                ],
                                Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-praktikanten-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"));
                            m.Zieldatei?.Erstellen(",", '\'', false, false);

                            m.Zieldatei = m.KlassenAnlegen(Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-klassen-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"));
                            m.Zieldatei?.Erstellen(",", '\'', false, false);

                            m.Schulpflichtüberwachung();

                            m.Zieldatei = m.GetFaecher(Path.Combine(Global.PfadExportdateien ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-faecher.csv"));
                            m.Zieldatei?.Erstellen(",", '\'', false, false);
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
                            m.FilterInteressierendeStudentsUndKlassen(configuration);

                            m.Zieldatei = m.Lernabschnittsdaten(Path.Combine(Global.PfadSchilddateien ?? "",
                                "SchuelerLernabschnittsdaten.dat"));
                            m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien,
                                ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"], []);
                            m.Zieldatei?.Erstellen("|", '\0', true, false);

                            m.Zieldatei = m.Leistungsdaten(Path.Combine(Global.PfadSchilddateien ?? "",
                                "SchuelerLeistungsdaten.dat"), configuration);
                            m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien,
                                ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt", "Fach"], ["Jahrgang"]);
                            m.Zieldatei?.Erstellen("|", '\0', true, false);

                            m.Zieldatei = m.Faecher(Path.Combine(Global.PfadSchilddateien ?? "", "Faecher.dat"));
                            m.Zieldatei?.Erstellen("|", '\0', true, false);

                            m.Zieldatei = m.Kurse(Path.Combine(Global.PfadSchilddateien ?? "", "Kurse.dat"));
                            m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, ["KursBez"],
                                ["Klasse", "Schulnr", "WochenstdPUNKTLEERZEICHENKL"]);
                            m.Zieldatei?.Erstellen("|", '\0', true, false);

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
                        "Mahnungen importieren (Gem. §50(4) SchulG)",
                        anrechnungen,
                        quelldateien.Notwendige(["marksperlesson", "schuelerleistungsdaten", "exportlessons", "studentgroupstudents", "schuelerleistungsdaten", "schuelerbasisdaten"]),
                        students,
                        Klassen,
                        [
                            "Die Datei SchuelerLeistungsdaten wird erstellt.",
                            "Die Datei SchuelerLeistungsdaten kann dann nach SchILD importiert werden.",
                        ],
                        m =>
                        {                            
                            m.Zieldatei = m.Leistungsdaten(Path.Combine(Global.PfadSchilddateien ?? "", "SchuelerLeistungsdaten.dat"), configuration, "Mahnung");
                            m.Zieldatei?.Erstellen("|", '\0', true, false);
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
                            m.FilterInteressierendeStudentsUndKlassen(configuration);

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
                            m.Zieldatei = m.Teilleistungen(@"ImportNachSchild\SchuelerTeilleistungen.dat");
                            m.Zieldatei.Erstellen("|", '\0', true, false);
                        }
                    ),
                    new Menüeintrag(
                        "Von PDF-Dateien in " + Global.PfadExportdateien + " verschlüsselte Kopien erstellen.",
                        anrechnungen,
                        new Dateien(),
                        students,
                        Klassen,
                        ["Von PDF-Dateien in " + Global.PfadExportdateien + " wird eine verschlüsselte Kopie erstellt.",
                        "Kopien bekommen die Dateiendung '-kennwort.pdf'"],
                        _ =>
                        {
                            var pdfDateien = new PdfDateien();
                            pdfDateien.KennwortSetzen(configuration);
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
                            m.Zieldatei = m.LernabschnittsdatenAlt(@"DatenaustauschSchild/SchuelerLernabschnittsdaten.dat", configuration);
                            m.Zieldatei.Erstellen("|", '\0', true, false);

                            m.Zieldatei = m.LeistungsdatenAlt(@"DatenaustauschSchild/SchuelerLeistungsdaten.dat");
                            m.Zieldatei?.Erstellen("|", '\0', true, false);
                        }
                    ),
                    new Menüeintrag(
                        "Fehlende Klassenbucheinträge bei den KuK anmahnen",
                        anrechnungen,
                        quelldateien.Notwendige(["lehrkraefte", "openperiod"]),
                        students,
                        Klassen,
                        [
                            "Die 10% der KuK mit den meisten offenen Klassenbucheinträgen werden angemahnt.",
                            "Die Anzahl der offenen Klassenbucheinträge wird aus der Datei 'OpenPeriods' ausgelesen.",
                            "Die KuK werden zuerst angezeigt. Vor dem Mailversand wird gefragt."
                        ],
                        m =>
                        {
                            lehrers = new Lehrers(m.Quelldateien);
                            lehrers.OffeneKlassenbuchEinträgeMahnen(m.Quelldateien);                            
                        }
                    ),
                    new Menüeintrag(
                        "PDF-Seiten nach E-Mail-Adressen von Lehrkräften durchsuchen und die betreffende Seiten mailen",
                        anrechnungen,
                        quelldateien.Notwendige(["lehrkraefte"]),
                        students,
                        Klassen,
                        [
                            "Die zuletzt bearbeitete PDF-Datei wird eingelesen.",
                            "Jede Seite der Datei wird nach E-Mail-Adressen durchsucht.",
                            "Die betreffenden Seiten werden an die E-Mail-Adressen gemailt.",
                            "Optional wird verschlüsselt."
                        ],
                        m =>
                        {
                            var pdfDatei = Directory.GetFiles(Global.PfadExportdateien, "*.pdf").OrderByDescending(File.GetLastWriteTime).FirstOrDefault();
                            Global.Konfig("PdfKennwort", configuration, "Kennwort für verschlüsselte PDFs angeben");
                            foreach (PdfSeite seite in (new PdfDatei(pdfDatei, new Lehrers(m.Quelldateien))).Seiten)
                            {
                                seite?.GetMailReceiver(lehrers);
                                seite?.PdfDocumentCreate(pdfDatei);
                                seite?.PdfDocumentEncrypt(Global.PdfKennwort);                                
                                seite?.Mailen("Nachricht aus der Schulverwaltung für", configuration);
                            }
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
                            m.Zieldatei = m.Zusatzdaten(@"ImportNachSchild/SchuelerZusatzdaten.dat");
                            m.Zieldatei.Erstellen("|", '\0', true, false);
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
                            m.Zieldatei = m.Basisdaten("ImportNachSchild/SchuelerBasisdaten.dat");
                            m.Zieldatei.Erstellen("|", '\0', true, false);
                        }
                    )
                ]
            );
        }
        catch
        {            
            return null;
        }   
    }
}
    