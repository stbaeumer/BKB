using Common;

public class Dateien : List<Datei>
{
    public Dateien()
    {
    }

    public void GetInteressierendeDateienMitAllenEigenschaften()
    {
        var schildhinweise = new string[]
        {
            "Exportieren Sie die Datei aus SchILD, indem Sie den Pfad gehen:",
            " Datenaustausch > Schnittstelle SchILD NRW > Export",
            " Alle oder einzelne Dateien auswählen.",
            " Die Datei(en) in " + Global.PfadExportdateien + " speichern."
        };

        var outlookhinweise = new string[]
        {
            "Exportieren Sie die Datei aus Outlook, indem Sie:",
            " ...",
            " ...",
            " ... "
        };

        var untishinweise = new string[]
        {
            "Exportieren Sie die Datei aus Untis, indem Sie den Pfad gehen:",
            " Datei > Import/Export > Export TXT Datei",
            " Die Datei auswählen.",
            " Die Datei in " + Global.PfadExportdateien + " speichern."
        };


        Add(new Datei(
            "SchildSchuelerExport",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus SchILD, indem Sie den Pfad gehen:",
                "Datenaustausch > Export in Text-/Exceldateien > Exportieren",
                "Die Vorlage 'SchildSchuelerExport' laden.",
                "Export starten.",
                "GGfs. muss die Vorlage erst erstellt werden. Es müssen folgende Felder enthalten sein:",
                "Geburtsdatum, Interne ID-Nummer, Nachname, Vorname, Klasse, Externe ID-Nummer, Status",
                "Delimiter: '|'; Dateiendung: *.dat"
            ],
            [""],
            true,
            d => d.FilternSchildSchuelerExport(),
            "*.txt"
        ));
        Add(new Datei(
            "Student_",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Webuntis, indem Sie als Administrator den Pfad gehen:",
                " Stammdaten > Schüler*innen > Berichte > Schüler > CSV-Ausgabe",
                " Die Datei in " + Global.PfadExportdateien + " speichern."
            ],
            [""],
            true,
            d => d.FilternWebuntisStudent(),
            "*.csv",
            "\t"
        ));        
        Add(new Datei(
            "GPU003",
            "Beschreibung",
            untishinweise,
            [""],
            false,
            d => d.FilternKlassenGPU003(),
            "*.TXT"
        ));        
        Add(new Datei(
            "Lehrkraefte.dat",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilterLehrkraefte()
        ));
        Add(new Datei(
            @"Gost.csv",
            "Beschreibung",
            [""],
            [""],
            false,
            d => d.FilterGost(),
            "*.csv",
            ","
        ));
        Add(new Datei(
            "ExportLessons",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilterExportLessons(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "SchuelerBasisdaten",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "Kurse.",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilterKurse()
        ));
        Add(new Datei(
            @"DatumsAusAtlantis.csv",
            "Beschreibung",
            [],
            [""],
            true,
            d => d.FilterDatumsAusAtlantis(),
            "*.csv",
            ","
        ));
        Add(new Datei(
            @"Adressen.csv",
            "Beschreibung",
            outlookhinweise,
            [""],
            true,
            d => d.FilternAdressenAtlantis(),
            "*.csv",
            ";"
        ));
        Add(new Datei(
            @"termine_kollegium.csv",
            "Beschreibung",
            outlookhinweise,
            [""],
            true,
            d => d.FilternTermineKollegium(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            @"termine_fhr.csv",
            "Beschreibung",
            outlookhinweise,
            [""],
            true,
            d => d.FilternTermineFhr(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            @"termine_verwaltung.csv",
            "Beschreibung",
            outlookhinweise,
            [""],
            true,
            d => d.FilternTermineVerwaltung(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            @"termine_berufliches_gymnasium.csv",
            "Beschreibung",
            outlookhinweise,
            [""],
            true,
            d => d.FilternTermineBeruflichesGymnasium(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            @"Atlantis-Zeugnisse-Noten.csv",
            "Beschreibung",
            [""],
            [""],
            false,
            d => d.FilternAtlantisZeugnisseNoten(),
            "*.csv",
            ","
        ));
        Add(new Datei(
            "SchuelerLeistungsdaten",
            "Beschreibung",
            schildhinweise,
            ["Mahnung", "Mahndatum", "Sortierung"],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerErzieher",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerTelefonnummern",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerAdressen",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "Klassen",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternSchildKlassen()
        ));
        Add(new Datei(
            "SchuelerZusatzdaten",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "Faecher",
            "Beschreibung",
            schildhinweise,
            ["Unterrichtsprache", "Sortierung S1", "Sortierung S2", "Gewichtung"],
            true,
            d => d.FilternSchildFaecher()
        ));
        Add(new Datei(
            "SchuelerLernabschnittsdaten",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerTeilleistungen",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerFehlstunden",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerVermerke",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "MarksPerLesson",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternMarksPerLessons(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "StudentgroupStudents",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternStudentgroupStudents(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "AbsencePerStudent",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternAbsencePerLEssons(),
            "*.csv",
            "\t"
        ));
    }

    public List<dynamic>? GetMatchingList(string pattern, Students students = null, Klassen klassen = null)
    {
        var datei = this.FirstOrDefault(datei => !string.IsNullOrEmpty(datei.Dateiname) && datei.Dateiname.ToLower().Contains(pattern, StringComparison.CurrentCultureIgnoreCase));
        
        if (datei == null)
        {            
            Global.ZeileSchreiben(pattern, "nicht vorhanden", ConsoleColor.DarkRed, ConsoleColor.White);    
            return [];

        }else if(datei.Count == 0)
        {
            if(datei.AbsoluterPfad.Length > 0)
            {
                Global.ZeileSchreiben(Path.GetFileName(datei.AbsoluterPfad) ?? "Unbekannt", datei.Count.ToString(), ConsoleColor.Red, ConsoleColor.White);
            }else{
                Global.ZeileSchreiben(pattern, "0", ConsoleColor.Red, ConsoleColor.White);
            }
        }
        else if (students == null && !datei.AbsoluterPfad.Contains("SchildSchuelerExport"))
        {            
            Global.ZeileSchreiben(Path.GetFileName(datei.AbsoluterPfad) + ": Schüler:", "0", ConsoleColor.Red,ConsoleColor.White);    
            return datei.ToList();
        }
        
        var gefilterteListe = datei.Filtern(students, klassen);

        if (datei.Count != 0 && datei.Erstelldatum.AddDays(Global.MaxDateiAlter) >= DateTime.Today)
        {
            // 2 Tage vor Ablauf wird die Farbe Gray  
            if (datei.Count != 0 && datei.Erstelldatum.AddDays(Global.MaxDateiAlter).AddDays(-2) >= DateTime.Today)
            {
                Global.ZeileSchreiben(  Path.GetFileName(datei.AbsoluterPfad) ?? "Unbekannt", gefilterteListe.Count.ToString(), ConsoleColor.Gray,ConsoleColor.White);    
            }
            else
            {
                Global.ZeileSchreiben(  Path.GetFileName(datei.AbsoluterPfad) ?? "Unbekannt", gefilterteListe.Count.ToString(), ConsoleColor.Yellow,ConsoleColor.White);
            }
                
            return gefilterteListe;
        }

        Console.ForegroundColor = ConsoleColor.Red;

        if (datei is { Count: 0 } && !datei.AbsoluterPfad.EndsWith(".dat"))
            Console.WriteLine("     Die Datei " + pattern + " enthält keine Elemente. Delimiter falsch?");
        if (datei is { Count: 0 } && datei.AbsoluterPfad.EndsWith(".dat"))
            Console.WriteLine("     Die Datei " + pattern + " enthält keine Elemente. Bitte prüfen.");
        if (datei != null && datei.Erstelldatum.AddDays(Global.MaxDateiAlter) <= DateTime.Today)
            Console.WriteLine("  Die Datei '" + pattern + "' darf höchstens " + Global.MaxDateiAlter +
                              " Tage alt sein. Erstelldatum: " + datei.Erstelldatum.ToShortDateString() +
                              ".\n  Sie ist veraltet und muss neu gezogen werden. \n  Dann neustarten.");
        Console.ResetColor();

        return [];
    }

    private string GetLastTwoPathParts(string path)
    {
        string directory = Path.GetDirectoryName(path) ?? "Unbekannt";
        string fileName = Path.GetFileName(path);

        if (directory == null) return fileName; // Falls kein Verzeichnis vorhanden ist

        string parentFolder = Path.GetFileName(directory); // Letzter Ordnername

        return $"{parentFolder}/{fileName}";
    }
    
    public Dateien Notwendige(List<string> dateinamenNotwendigeDateien)
    {
        Dateien notwendige = new Dateien();

        foreach (var dateinameNotwendig in dateinamenNotwendigeDateien)
        {
            try
            {  
                var datei = this.First(datei => !string.IsNullOrEmpty(datei.Dateiname) && datei.Dateiname.ToLower().StartsWith(dateinameNotwendig.ToLower(), StringComparison.CurrentCultureIgnoreCase));
    
                var absoluterPfad = this.First(datei => !string.IsNullOrEmpty(datei.Dateiname) && datei.Dateiname.ToLower().StartsWith(dateinameNotwendig.ToLower(), StringComparison.CurrentCultureIgnoreCase)).AbsoluterPfad; 

                if(absoluterPfad.Length > 0)
                {                          
                    if(datei.Count == 0){
                        datei.Fehlermeldung = absoluterPfad + " existiert, ist aber leer.";
                    }else{
                        if(datei.Erstelldatum.AddDays(Global.MaxDateiAlter) <= DateTime.Now){

                            datei.Fehlermeldung = absoluterPfad + " existiert, ist aber veraltet.";
                        }
                    }  
                }
                else
                {
                    datei.Fehlermeldung = Global.PfadExportdateien + ": Die Datei '" + dateinameNotwendig + "' existiert nicht.";                 
                }
                notwendige.Add(datei);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);                
            }
        }

        return notwendige;
    }

    public void GetZeilen(List<string> dateienImPfad = null)
    {
        Global.Veraltet = false;

        if (dateienImPfad == null)
        {
            dateienImPfad = GetDateienImPfad();    
        }

        // Für jede interessierende Quelldatei
        foreach (var datei in this)
        {
            // ... sofern es eine passende Datei gibt ... 
            if (dateienImPfad.Any(d => Path.GetFileName(d).ToLower().StartsWith(datei.Dateiname.ToLower())))            
            {
                datei.AbsoluterPfad = dateienImPfad.OrderBy(File.GetCreationTime).LastOrDefault(d => Path.GetFileName(d).StartsWith(datei.Dateiname));
                datei.Erstelldatum = File.GetCreationTime(datei.AbsoluterPfad);
                
                datei.GetZeilen();

                datei.ZeileSchreiben();
                
                if (datei.Erstelldatum.AddDays(Global.MaxDateiAlter) <= DateTime.Now)
                {
                    Global.Zeilen.Add(new ValueTuple<string, ConsoleColor>("Mindestens ein der hochgeladenen Dateien ist zu alt.", ConsoleColor.Red));                 
                    Global.Veraltet = true;
                }
            }
        }
        if (Global.Veraltet)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("                                                               *) Veraltete Dateien.");
            Console.ResetColor();
        }
    }

    public List<string> GetKlassen()
    {
        var klassenSet = new HashSet<string>(); // HashSet für einzigartige Werte

        foreach (var datei in this.Where(x => x.Count > 0))
        {
            foreach (var record in datei) // Direkt durch die bereits geladenen Records iterieren
            {
                var dict = (IDictionary<string, object>)record; // FastDynamicObject als Dictionary behandeln

                // Suche nach einer passenden Spalte
                var key = dict.Keys.FirstOrDefault(k =>
                    k.Equals("Klasse", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("klasse", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Klassen", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("klassen", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Class", StringComparison.OrdinalIgnoreCase));

                if (key != null && dict[key] != null)
                {
                    var value = dict[key].ToString().Trim();

                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        // Falls mehrere Klassen mit "~" getrennt sind, aufsplitten
                        var klassenArray = value.Split('~', StringSplitOptions.RemoveEmptyEntries);

                        foreach (var klasse in klassenArray)
                        {
                            klassenSet.Add(klasse.Trim()); // Trim, um unnötige Leerzeichen zu entfernen
                        }
                    }
                }
            }
        }

        // Sortierte Liste erstellen und "alle" an den Anfang setzen
        var klassenListe = klassenSet.Where(x => !string.IsNullOrWhiteSpace(x))
            .OrderBy(x => x)
            .ToList();

        klassenListe.Insert(0, "Für alle Klassen oder "); // "alle" an erster Stelle einfügen

        return klassenListe;
    }

    public List<string> GetDateienImPfad()
    {
        return Directory.GetFiles(Global.PfadExportdateien, "*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) ||
                f.EndsWith(".dat", StringComparison.OrdinalIgnoreCase) ||
                f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => File.GetCreationTime(f)) // Sortierung nach Erstellungsdatum
                .ToList();
    }
}