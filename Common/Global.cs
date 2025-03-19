using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.Json;
using Common;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using MySqlConnector;
using Path = System.IO.Path;

public static class Global
{
    public static List<(string Meldung, ConsoleColor Farbe)> Zeilen = new();
    public static int Kalenderwoche { get; set; }
    public static string? AusbuchenNachWievielTagen { get; set; }
    public static string? WikiJsonUserKennwort { get; set; }

    public static string? PfadSchilddateien { get; set; }
    public static string? WikiJsonUser { get; set; }
    public static string? WikiUrl { get; set; }
    public static string? MailDomain { get; set; }
    public static string? MariaUser { get; set; }
    public static string? MariaIp { get; set; }
    public static string? MariaPort { get; set; }
    public static string? MariaPw { get; set; }
    public static string? PdfKennwort { get; set; }

public static string? NetmanMailReceiver { get; set; }

public static string? SmtpServer { get; set; }
public static string? SmtpPort { get; set; }
public static string? SmtpPassword { get; set; }
public static string? SmtpUser { get; set; }

    public static DateTime Sprechtagsdatum { get; set; }
    public static DateTime Halbjahreskonferenzdatum { get; set; }
    public static DateTime Halbjahreszeugnisdatum { get; set; }
    public static DateTime Jahreskonferenzdatum { get; set; }
    public static DateTime Jahreszeugnisdatum { get; set; }
    public static string? AktuellerPfad { get; set; }
    public static string? InputFolder { get; set; }
    public static string? OutputFolder { get; set; }
    public static string? MariaDb { get; set; }
    public static string? User { get; set; }
    public static int MaxDateiAlter { get; set; }
    public static int MaximaleAnzahlFehlstundenProTag { get; set; }
    public static int FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt { get; set; }
    public static DateTime DatenimportLetztesDatum { get; set; }
    public static object? WikiSprechtagKleineAenderung { get; set; }
    public static List<string>? Protokoll { get; set; }
    public static Dateien? Dateien { get; internal set; }
    public static string PfadExportdateien { get; set; }
    public static string Abschnitt { get; set; }

    public enum Datentyp
    {
        String,
        Int,
        DateTime,
        Pfad,
        Url,
        JaNein
    }

    public enum Header
    {
        H1,
        H2,
        H3
    };

    public enum Protokollieren
    {
        Ja,
        Nein
    };


    public static List<string> AktSj = new List<string>()
    {
        (DateTime.Now.Month >= 7 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
        (DateTime.Now.Month >= 7 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString()
    };

    public static string ConnectionStringUntis { get; set; }
    public static string H1 { get; set; }
    public static string ZipKennwort { get; set; }
    public static Boolean CodeSpace { get; set; }
    public static bool Veraltet { get; internal set; }
    public static ConsoleColor DefaultBackground { get; set; }

    public static string? SafeGetString(SqlDataReader reader, int colIndex)
    {
        if (!reader.IsDBNull(colIndex))
            return reader.GetString(colIndex);
        return string.Empty;
    }

    public static string SafeGetString(MySqlDataReader reader, int colIndex)
    {
        if (!reader.IsDBNull(colIndex))
            return reader.GetString(colIndex);
        return string.Empty;
    }

    internal static void Dateischreiben(string name)
    {
        string tempPfadUndDatei = Path.Combine(Path.GetTempPath(), Path.GetFileName(name));
        string pfadUndDatei = Path.Combine(Global.PfadExportdateien, name);
        UTF8Encoding utf8NoBom = new UTF8Encoding(false);

        if (File.Exists(pfadUndDatei) && File.Exists(tempPfadUndDatei))
        {
            string contentNeu = File.ReadAllText(tempPfadUndDatei, utf8NoBom);

            // Lese den Inhalt der Dateien
            string contentAlt = File.ReadAllText(pfadUndDatei, utf8NoBom);

            // Vergleiche die Inhalte der Dateien
            if (contentAlt != contentNeu)
            {
                // Überschreibe alt mit dem Inhalt von neu
                File.WriteAllText(pfadUndDatei, contentNeu, utf8NoBom);
                Global.ZeileSchreiben(name, "überschrieben", ConsoleColor.Yellow,ConsoleColor.Gray);
            }
            else
            {
                Global.ZeileSchreiben(name, "Identisch. Keine Änderungen", ConsoleColor.White, ConsoleColor.Black);
            }
        }

        if (!File.Exists(pfadUndDatei))
        {
            string directoryPath = Path.GetDirectoryName(pfadUndDatei);

            if (directoryPath != null)
            {
                // Fehlende Verzeichnisse anlegen
                Directory.CreateDirectory(directoryPath);
            }

            string contentNeu = File.ReadAllText(tempPfadUndDatei, utf8NoBom);
            File.WriteAllText(pfadUndDatei, contentNeu, utf8NoBom);
            Global.ZeileSchreiben(name, ": Datei neu erstellt.", ConsoleColor.Green, ConsoleColor.Gray);
        }
    }

    public static void DisplayHeader(Header header, string text, Protokollieren protokollieren)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;

        if (header == Header.H1)
        {
            Console.Clear();
            int platzVorher = (Console.WindowWidth - text.Length) / 2;
            Console.WriteLine(text.PadLeft(platzVorher + text.Length));
            
            
            // Zweite Zeile
            Console.BackgroundColor = ConsoleColor.Yellow;
            Console.ForegroundColor = ConsoleColor.Black;
            text = "";
            platzVorher = (Console.WindowWidth - text.Length) / 2;
            Console.WriteLine((text.PadLeft(platzVorher + text.Length)).PadRight(platzVorher + text.Length + platzVorher));
            Console.ResetColor();
            Console.ResetColor();
            //Console.WriteLine("".PadRight(Console.WindowWidth, '='));

            if (protokollieren == Protokollieren.Ja)
            {
                Protokoll.Add(text.PadLeft(platzVorher + text.Length));
                Protokoll.Add("".PadRight(Console.WindowWidth, '='));
                Protokoll.Add("");
            }
        }

        if (header == Header.H2)
        {
            Console.WriteLine("");
            int platzVorher = (Console.WindowWidth - text.Length) / 2;
            //Console.WriteLine(text.PadLeft(platzVorher + text.Length));
            Console.WriteLine("  " + text + ":");
            //Console.WriteLine("".PadRight(Console.WindowWidth, '-'));
            Console.WriteLine("");

            if (protokollieren == Protokollieren.Ja)
            {
                Protokoll.Add("");
                Protokoll.Add(text.PadLeft(platzVorher + text.Length));
                Protokoll.Add("".PadRight(Console.WindowWidth, '-'));
                Protokoll.Add("");
            }
        }

        if (header == Header.H3)
        {
            Console.WriteLine("");
            Console.WriteLine("   " + text);
            Console.WriteLine("   " + "".PadRight(text.Length, '*'));

            if (protokollieren == Protokollieren.Ja)
            {
                Protokoll.Add("");
                Protokoll.Add(" " + text);
                Protokoll.Add(" " + "".PadRight(text.Length, '*'));
            }
        }

        Console.ForegroundColor = ConsoleColor.White;
    }

    internal static Datei GetDat(string Pfad)
    {
        return Global.Dateien
            .Where(d => d.UnterordnerUndDateiname != null)
            .FirstOrDefault(d => d.UnterordnerUndDateiname.ToLower()
                .Contains(Pfad.ToLower()));
    }

    public static string InsertLineBreaks(string text, int maxLineLength)
    {
        if (string.IsNullOrEmpty(text) || maxLineLength <= 0)
            return text;

        var currentIndex = 0;
        var length = text.Length;
        var result = new StringBuilder();

        while (currentIndex < length)
        {
            // Calculate the length of the next segment
            int nextSegmentLength = Math.Min(maxLineLength, length - currentIndex);
            // Append the segment and a line break
            result.Append(text.Substring(currentIndex, nextSegmentLength));
            result.Append(Environment.NewLine + "   ");
            // Move to the next segment
            currentIndex += nextSegmentLength;
        }

        return result.ToString();
    }

    internal static void OrdnerAnlegen(string datei)
    {
        string temp = Path.GetTempPath();

        if (!Directory.Exists(temp))
        {
            Directory.CreateDirectory(temp!);
            Console.WriteLine($"Verzeichnis erstellt: {temp}");
        }

        var verzeichnis = Path.GetDirectoryName(Path.Combine(Global.PfadExportdateien, datei));

        if (Directory.Exists(verzeichnis)) return;
        Directory.CreateDirectory(verzeichnis!);
        Console.WriteLine($"Verzeichnis erstellt: {verzeichnis}");
    }

    internal static void OrdnerAnlegen(object name)
    {
        throw new NotImplementedException();
    }

    public static void ZeileSchreiben(string linkeSeite, string rechteSeite, ConsoleColor foreground = ConsoleColor.Black, ConsoleColor background = ConsoleColor.White)
    {             
        var gesamtbreite = Console.WindowWidth;
        var punkte = gesamtbreite - linkeSeite.Length - rechteSeite.Length - 6;
        var mitte = " .".PadRight(Math.Max(3, punkte), '.') + " ";

        // Wenn linkeSeite auf einen Punkt endet, dann wird das Leerzeichen links durch einen Punkt ersetzt

        if (linkeSeite.Substring(linkeSeite.Length - 1, 1) == ".")
        {
            mitte = "." + mitte.Substring(1);
        }

        WriteLine("".PadRight(2) + linkeSeite + mitte + rechteSeite + " ", foreground, background);
    }
    
    private static void WriteLine(string meldung, ConsoleColor foreground, ConsoleColor background)
    {
        Console.ResetColor();
        Console.BackgroundColor = background;
        Console.ForegroundColor = foreground;        

        // Schwarze der weiße Hintergründe werden durch default ersetzt
        if (background == ConsoleColor.White || background == ConsoleColor.Black){
            Console.BackgroundColor = Global.DefaultBackground;
        }
        
        Console.WriteLine(meldung);
        Console.ResetColor();
    }

    internal static void DisplayCenteredBox(List<string> text, int prozent)
    {
        // Breite der Konsole und die Box-Breite berechnen
        int consoleWidth = Console.WindowWidth > 0 ? Console.WindowWidth : 125;
        int boxWidth = consoleWidth * prozent / 100;

        // Text in Zeilen aufteilen, falls \n vorhanden ist
        string[] lines = text.ToArray();

        var wrappedLines = new List<string>();

        foreach (var line in lines)
        {
            // Text passend umbrechen
            string[] words = line.Split(' ');
            string currentLine = "";

            foreach (var word in words)
            {
                if (currentLine.Length + word.Length + 1 > boxWidth - 4) // 4 Zeichen für die Ränder abziehen
                {
                    wrappedLines.Add(currentLine);
                    currentLine = word;
                }
                else
                {
                    if (currentLine.Length > 0)
                        currentLine += " ";
                    currentLine += word;
                }
            }

            if (currentLine.Length > 0)
            {
                wrappedLines.Add(currentLine);
            }
        }

        // Zentrierung berechnen
        int leftPadding = (consoleWidth - boxWidth) / 2;

        // Horizontale Grenze für die Box
        string horizontalBorder = new string('─', boxWidth - 2);

        // Box zeichnen
        // Oberer Rand mit Padding
        Console.WriteLine(new string(' ', leftPadding) + "┌" + horizontalBorder + "┐");

        // Textzeilen in der Box
        foreach (var wrappedLine in wrappedLines)
        {
            // Leerzeichen auffüllen, damit die Zeile genau so breit ist wie die Box
            string paddedLine = wrappedLine.PadRight(boxWidth - 2);
            Console.WriteLine(new string(' ', leftPadding) + "│" + paddedLine + "│");
        }

        // Unterer Rand mit Padding
        Console.WriteLine(new string(' ', leftPadding) + "└" + horizontalBorder + "┘");
    }

    public static void Speichern(string key, string value)
{
    var documentsFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
    var configPath = Path.Combine(documentsFolderPath, "BKB.json");
                
    // Aktuelle JSON-Daten lesen
    string jsonFilePath = configPath;
    var json = File.ReadAllText(jsonFilePath);
    var jsonDoc = JsonDocument.Parse(json);
    var jsonRoot = jsonDoc.RootElement;
    
    string finalValue = Verschluesseln(value);

    // Neuen Wert setzen
    using (var stream = new MemoryStream())
    {
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true }))
        {
            writer.WriteStartObject();
            foreach (var property in jsonRoot.EnumerateObject())
            {
                if (property.NameEquals(key))
                {
                    writer.WriteString(key, finalValue);
                }
                else
                {
                    property.WriteTo(writer);
                }
            }

            writer.WriteEndObject();
        }

        // Neue JSON-Daten in die Datei schreiben
        File.WriteAllText(jsonFilePath, System.Text.Encoding.UTF8.GetString(stream.ToArray()));
    }
}

// Hilfsmethode zur Verschlüsselung
    public static string Verschluesseln(string value)
    {
        // Beispiel für eine einfache Verschlüsselung (Base64)
        byte[] data = Encoding.UTF8.GetBytes(value);
        string encryptedValue = Convert.ToBase64String(data);
        return encryptedValue;
    }
    public static void OpenCurrentFolder()
    {
        Console.WriteLine("     Der Ordner " + Environment.CurrentDirectory + " wird geöffent.");
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = Environment.CurrentDirectory,
                UseShellExecute = true,
                Verb = "open"
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine("    Fehler beim Öffnen des Ordners:");
            Console.WriteLine(ex.Message);
        }
    }

    public static void OpenWebseite(string url)
    {
        if (!url.StartsWith("http"))
        {
            url = "https://wiki.svws.nrw.de/mediawiki/index.php?title=Schnittstellenbeschreibung#";
            Console.WriteLine("     Die Seite " + url + " wird geöffnet.");
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url.TrimEnd('#'),
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine("    Fehler beim Öffnen der Webseite:");
            Console.WriteLine(ex.Message);
        }
    }

    internal static void ProtokollzeileSchreiben(string v, string protokollDatei)
    {
        using (StreamWriter writer = new StreamWriter(protokollDatei))
        {
            writer.WriteLine(v);
        }
    }

    public static string CheckFile(string Pfad, string endung)
    {
        Pfad = Path.Combine(Global.PfadExportdateien, Pfad);

        if (!Path.Exists(Path.GetDirectoryName(Pfad)))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Pfad));
        }

        var sourceFile =
            (from f in Directory.GetFiles(Path.GetDirectoryName(Pfad), endung, SearchOption.AllDirectories)
                where Path.GetFileName(f).StartsWith(Path.GetFileName(Pfad))
                orderby File.GetLastWriteTime(f)
                select f).LastOrDefault();

        return sourceFile;
    }

    internal static List<object> ZeileSchreiben(List<object> absencePerStudents)
    {
        Global.ZeileSchreiben(absencePerStudents.ToString(), absencePerStudents.Count().ToString(), ConsoleColor.DarkMagenta, ConsoleColor.White);
        return absencePerStudents;
    }

    public static class GlobalCsvMappings
    {
        public static Dictionary<string, List<string>> AliasMappings = new Dictionary<string, List<string>>();

        // Diese Methode liest die Alias.csv ein
        public static void LoadMappingsFromFile(string aliasFilePath)
        {
            using (var reader = new StreamReader(aliasFilePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<CsvAlias>();
                foreach (var record in records)
                {
                    // Jede PropertyName wird mit allen entsprechenden CSV-Namen verknüpft
                    AliasMappings[record.PropertyName] = new List<string> { record.CsvName1, record.CsvName2 };
                }
            }
        }

        public static void AddMappings<T>(ClassMap<T> map)
        {
            var properties = typeof(T).GetProperties(); // Hole alle Eigenschaften der Klasse T

            foreach (var property in properties)
            {
                var propertyName = property.Name;

                // Prüfe, ob die Eigenschaft in der Alias.csv vorhanden ist
                if (AliasMappings.TryGetValue(propertyName, out var csvNames))
                {
                    // Wenn ein Mapping in Alias.csv gefunden wurde, verwende die dynamischen Namen
                    var mapExpression = GetPropertyMappingExpression<T>(propertyName);
                    if (mapExpression != null)
                    {
                        map.Map(mapExpression).Name(csvNames.ToArray());
                    }
                }
                else
                {
                    // Fallback: Verwende den Standardnamen der Eigenschaft, wenn kein Mapping gefunden wurde
                    map.Map(typeof(T), property).Name(propertyName);
                }
            }
        }



        // Dynamische Methode, um eine Eigenschaft basierend auf dem Namen zu mappen
        private static System.Linq.Expressions.Expression<Func<T, object>> GetPropertyMappingExpression<T>(
            string propertyName)
        {
            var param = System.Linq.Expressions.Expression.Parameter(typeof(T), "m");
            var property = typeof(T).GetProperty(propertyName);
            if (property == null) return null;

            var body = System.Linq.Expressions.Expression.Convert(
                System.Linq.Expressions.Expression.Property(param, property),
                typeof(object)
            );
            return System.Linq.Expressions.Expression.Lambda<Func<T, object>>(body, param);
        }
    }

    public class CsvAlias
    {
        public required string PropertyName { get; set; }
        public required string CsvName1 { get; set; }
        public required string CsvName2 { get; set; }
    }

    internal static string ListeErzeugen(List<string> kategorien, char delimiter)
    {
        var x = "";

        foreach (var item in kategorien)
        {
            x += item.Trim() + delimiter;
        }

        return x.TrimEnd(delimiter);
    }

    public static void Konfig(string parameter, IConfiguration configuration, string beschreibung = "", Datentyp datentyp = Datentyp.String)
    {
        beschreibung = string.IsNullOrEmpty(beschreibung) ? "Bitte " + parameter + " eingeben " : beschreibung;
        
        var value = Entschluesseln(configuration[parameter]);
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
                if (true)
                {
                    string currentValueStr = !string.IsNullOrEmpty(userInput) ? currentValue.ToString() : "";

                    Console.ForegroundColor = ConsoleColor.Green;

                    var wert = currentValueStr.Replace(" 00:00:00", "");
                    if (parameter.Contains("ennwort") || parameter.Contains("asswor"))
                    {
                        wert = wert.Substring(0, Math.Min(wert.Length, 3)) + "**********";
                    }

                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.Write($"  {beschreibung} [{wert}] : ");
                    
                    userInput = Console.ReadLine();

                    // Wenn Enter gedrückt wird, bleibt der aktuelle Wert erhalten

                    if (string.IsNullOrEmpty(userInput))
                    {
                        //Console.WriteLine($"  {parameter} bleibt unverändert: {currentValueStr}");
                        return;
                    }

                    if (userInput == "q")
                    {
                        return;
                    }

                    if (userInput == "ö")
                    {
                        //todo:
                        Global.OpenCurrentFolder();
                    }

                    if (userInput == "y")
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
                                Console.WriteLine("   Ungültige Zahl. Versuchen Sie es erneut.");
                            firstrun = false;                            
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
                                Console.WriteLine("    Ungültige Eingabe für DateTime. Versuchen Sie es erneut.");
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
                                    "    Ungültige EIngabe. Erlaubt sind: j und n. Versuchen Sie es erneut.");
                            firstrun = false;
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
                                Console.WriteLine("    Ungültiger Pfad. Versuchen Sie es erneut.");
                            firstrun = false;
                            continue;
                        }

                        break;

                    default:
                        Console.WriteLine("    Unbekannter Datentyp.");
                        continue;
                }
                Console.ForegroundColor = ConsoleColor.Cyan;
                // Falls der Wert erfolgreich geparst wurde, setzen wir ihn und beenden die Schleife
                if (convertedValue == null) continue;
                Speichern(parameter, userInput);
                property.SetValue(null, convertedValue);
                Console.WriteLine("  " +parameter.PadRight(41) + convertedValue.ToString().Replace(" 00:00:00", ""));
                break;
            }
        }
        else
        {
            Console.WriteLine($"Eigenschaft '{parameter}' nicht gefunden oder nicht beschreibbar.");
        }
    }

    static bool IsValidHttpUrl(string url)
    {
        return Uri.TryCreate(url, UriKind.Absolute, out Uri? uriResult) &&
               (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
    }


    public static void GetKalenderwoche()
    {
        // Example date
        var date = DateTime.Now;

        // Define the ISO 8601 calendar
        var ci = CultureInfo.CurrentCulture;
        var calendar = ci.Calendar;

        // Define the CalendarWeekRule and the DayOfWeek for the first day of the week
        const CalendarWeekRule weekRule = CalendarWeekRule.FirstFourDayWeek;
        const DayOfWeek firstDayOfWeek = DayOfWeek.Monday;

        // Get the week number for the specified date
        Kalenderwoche = calendar.GetWeekOfYear(date, weekRule, firstDayOfWeek);
    }

    public static void DateiErstellen(List<dynamic>? neue, string datei, string delimiter, char quote, bool encoding,
        bool shouldAllQuote,
        Dateien dateien)
    {
        if (neue.Count == 0)
        {
            return;
        }

        Exception fehler = null!;
        var pfad = Path.Combine(Global.PfadExportdateien, datei);

        Global.AktuellerPfad = Path.GetDirectoryName(pfad);

        try
        {
            if (!Path.Exists(Path.GetDirectoryName(pfad)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(pfad)!);
            }

            var config = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter,
                Quote = quote,
                ShouldQuote = args => shouldAllQuote
            };

            File.Delete(pfad);

            if (neue != null && neue.Any())
            {
                using var writer = new StreamWriter(pfad, false, new System.Text.UTF8Encoding(encoding));
                using var csv = new CsvWriter(writer, config);

                // Header manuell extrahieren
                var firstRecord = neue[0] as IDictionary<string, object>;
                var adjustedHeaders = new List<string>();

                foreach (var header in firstRecord.Keys)
                {
                    // Anpassen der Header
                    var adjustedHeader = header
                        .Replace("DOPPELPUNKT", ":")
                        .Replace("PUNKT", ".")                        
                        .Replace("MINUS", "-")
                        .Replace("ZWEI", "2")
                        .Replace("EINS", "1")
                        .Replace("UNTERSTRICH", "_")
                        .Replace("SLASH", "/")
                        .Replace("LEERZEICHEN", " ");

                    adjustedHeaders.Add(adjustedHeader);
                }

                // Schreiben der angepassten Header
                foreach (var header in adjustedHeaders)
                {
                    csv.WriteField(header);
                }

                csv.NextRecord();

                // Schreiben der Datensätze
                foreach (var record in neue)
                {
                    var recordDict = record as IDictionary<string, object>;
                    foreach (var value in recordDict.Values)
                    {
                        csv.WriteField(value);
                    }

                    csv.NextRecord();
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Fehler beim Schreiben der Datei: {ex.Message}");
            fehler = ex;
        }
        finally
        {
            Global.ZeileSchreiben(pfad + " - Zeilen:", neue.Count().ToString(), ConsoleColor.DarkYellow, ConsoleColor.DarkGreen);

            var rechteSeite = "bereit für den Import nach ";

            if (PfadExportdateien.ToLower().Contains("schild"))
            {
                rechteSeite += "SchILD";
            }

            if (PfadExportdateien.ToLower().Contains("webuntis"))
            {
                rechteSeite += "Webuntis";
            }

            if (PfadExportdateien.ToLower().Contains("geevoo"))
            {
                rechteSeite += "Geevoo";
            }



            Global.ZeileSchreiben(pfad + " neu erstellt", rechteSeite, ConsoleColor.Yellow,ConsoleColor.Gray);
            Console.ForegroundColor = ConsoleColor.White;
        }
    }


    private static void Zeilenconfigurationn(string datei, List<dynamic>? neue, List<dynamic> vorhandene,
        string[] schluesselIgnorieren)
    {
        foreach (var neu in neue)
        {
            var neuDict = (IDictionary<string, object>)neu;

            // Finde die passende Zeile in 'vorhandene'
            var match = vorhandene.FirstOrDefault(vorhanden =>
            {
                var vorhandenDict = (IDictionary<string, object>)vorhanden;

                // Vergleiche nur relevante Schlüssel
                return neuDict.Keys
                    .Where(key => !schluesselIgnorieren.Contains(key))
                    .All(key =>
                        vorhandenDict.ContainsKey(key) &&
                        neuDict[key]?.ToString() == vorhandenDict[key]?.ToString());
            });

            // Wenn keine Übereinstimmung gefunden wurde, gibt es nichts zu aktualisieren
            if (match == null) continue;

            var matchDict = (IDictionary<string, object>)match;

            // Prüfe auf Unterschiede und protokolliere configurations
            bool needsconfiguration = false;
            Console.WriteLine("Zeile, die aktualisiert werden muss:");
            foreach (var key in neuDict.Keys.Where(key => !schluesselIgnorieren.Contains(key)))
            {
                if (!matchDict.ContainsKey(key) || neuDict[key]?.ToString() != matchDict[key]?.ToString())
                {
                    Console.WriteLine($"{key}: {matchDict[key]} --> {neuDict[key]}");
                    needsconfiguration = true;
                }
            }

            if (needsconfiguration)
            {
                Console.WriteLine("----------------------");
            }
        }
    }

    public static void EditorOeffnen(string pfad)
    {
        try
        {
            System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Notepad++\Notepad++.exe", pfad);
        }
        catch (Exception)
        {
            System.Diagnostics.Process.Start("Notepad.exe", pfad);
        }
    }

    /// <summary>
    /// Die Exportdateien aus SchILD werden zu den anderen heruntergeladenen Dateien verschoben, um Platz zu machen für neue
    /// Dateien, die dann nach SchILD importiert werden. 
    /// </summary>
    public static void ExportAusSchildVerschieben(IConfiguration configuration)
    {
        // Stelle sicher, dass der Zielordner existiert
        if (!Directory.Exists(PfadSchilddateien))
        {
            if (PfadSchilddateien != null) Directory.CreateDirectory(PfadSchilddateien);
        }

        // Hole alle .dat-Dateien aus dem Quellordner
        if (PfadSchilddateien == null) return;

        if (!Directory.Exists(PfadSchilddateien))
        {            
           Konfig("PfadSchilddateien", configuration, "Bitte geben Sie den Pfad für die Schilddateien an:", Datentyp.Pfad);
        }

        // Die SchildSchuelerExport wird immer kopiert.

        var datei = Directory.GetFiles(Global.PfadSchilddateien, "*", SearchOption.TopDirectoryOnly).FirstOrDefault(f => Path.GetFileName(f).ToLower().Contains("schildschuelerexport"));

        if (datei != null)
        {
            var destinationPath = Path.Combine(Global.PfadExportdateien, Path.GetFileName(datei));
            File.Copy(datei, destinationPath, true);
            Console.WriteLine("Die Datei " + Path.GetFileName(datei) + " wurde erfolgreich nach " + Global.PfadExportdateien + " kopiert.");
        }

        // Hole alle .dat-Dateien aus dem Quellordner
        if (PfadSchilddateien == null) return;

        var datFiles = Directory.GetFiles(PfadSchilddateien, "*.dat").ToList();

        if (datFiles.Count > 5)
        {
            // Hole die Erstellungszeiten der Dateien
            var creationTimes = datFiles
                .Select(file => File.GetCreationTime(file))
                .OrderBy(time => time)
                .ToList();

            // Prüfe, ob alle Dateien innerhalb von einer Minute erstellt wurden
            var withinOneMinute = creationTimes.Last() - creationTimes.First() <= TimeSpan.FromMinutes(4);

            if (withinOneMinute)
            {
                var existingDatFiles = Directory.GetFiles(PfadExportdateien, "*.dat");
                foreach (var file in existingDatFiles)
                {
                    File.Delete(file);
                }

                // Verschiebe die Dateien
                foreach (var file in datFiles)
                {
                    var destinationPath = Path.Combine(PfadExportdateien, Path.GetFileName(file));
                    File.Move(file, destinationPath);
                }

                Console.WriteLine("Alle ExportDateien aus Schild wurden erfolgreich nach " + PfadExportdateien +
                                  " verschoben.");
            }
            else
            {
                Console.WriteLine("Die Dateien wurden nicht innerhalb von einer Minute erstellt.");
            }
        }
        else
        {
            foreach (var file in datFiles)
            {
                File.Delete(file);
            }
        }
    }

    public static string PrüfeAufNullOrEmpty(IDictionary<string, object> dict, string s)
    {
        if (dict.TryGetValue(s, out var nameObj) && nameObj is string name && !string.IsNullOrWhiteSpace(name))
        {
            return name;
        }
        else
        {
            return "";
        }
    }

    public static List<string> FindeNeuesteDateien(List<string> suchWoerter)
    {
        var ergebnisse = new List<string>();

        try
        {
            var alleDateien = Directory.GetFiles(Global.PfadExportdateien, "*", SearchOption.AllDirectories)
                .Select(datei => new FileInfo(datei));

            foreach (var wort in suchWoerter)
            {
                var neuesteDatei = alleDateien
                    .Where(datei => datei.Name.Contains(wort, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(datei => datei.LastWriteTime) // Sortiere nach Änderungsdatum absteigend
                    .FirstOrDefault();

                if (neuesteDatei != null)
                {
                    ergebnisse.Add(neuesteDatei.FullName);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fehler: " + ex.Message);
        }

        return ergebnisse;
    }

    public static void EinstellungenDurchlaufen(IConfiguration configuration)
    {        
        ZeileSchreiben("Einstellungen", PfadSchilddateien, ConsoleColor.Black, ConsoleColor.Cyan);
        
        Konfig("PfadExportdateien", configuration, @"Downloads-Ordner angeben", Datentyp.Pfad);
        Konfig("PfadSchilddateien", configuration, @"SchILD-Dateien-Ordner", Datentyp.Pfad);
        Konfig("MaxDateiAlter", configuration, "Wie viele Tage dürfen Dateien höchstens alt sein?", Datentyp.Int);
        Konfig("ZipKennwort", configuration, "Kennwort zum Verschlüsseln von Zip-Dateien");
        Konfig("ConnectionStringUntis", configuration, "ConnectionStringUntis eingeben (optional)");
        Konfig("MailDomain", configuration, "Mail-Domain für Schüler*innen eingeben");
        Konfig("SmtpUser", configuration, "Mail-Benutzer angeben (optional)");
        Konfig("SmtpPassword", configuration, "Mail-Kennwort eingeben (optional)");
        Konfig("SmtpPort", configuration, "SMTP-Port eingeben (optional)");
        Konfig("SmtpServer", configuration, "SMTP-Server angeben (optional)");
        Konfig("NetmanMailReceiver", configuration, "Wem soll die Netman-Mail geschickt werden");
        
        Speichern("EinstellungenVorgenommen", configuration["EinstellungenVorgenommen"] = "j");
    }

    public static bool RunningInCodeSpace()
    {
        if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("CODESPACES")))
        {
            return true;
        }        
        else
        {            
            return false;
        }
    }

    public static void WeiterMitAnykey(IConfiguration configuration)
    {                    
        if(!Global.CodeSpace)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = Global.PfadSchilddateien,
                UseShellExecute = true
            });
        }

        Console.ResetColor();
        Console.WriteLine(" ");
        ZeileSchreiben("Mit Anykey zurück zum Menü", "oder: x : Einstellungen öffnen; y : Dokumentation öffnen", ConsoleColor.Black, ConsoleColor.Gray);
       
        var weiter = Console.ReadKey(true); // true unterdrückt die Ausgabe des Zeichens im Terminal

        if (weiter.Key == ConsoleKey.X)
        {
            EinstellungenDurchlaufen(configuration);
            return;
        }
    
        if (weiter.Key == ConsoleKey.Y)
        {
            OpenWebseite("https://github.com/stbaeumer/BKB/wiki");
            return;
        }

        Console.ResetColor();
        Console.Clear();
    }
    public static string Entschluesseln(string encryptedValue)
    {
        // Beispiel für eine einfache Entschlüsselung (Base64)
        try
        {
            byte[] data = Convert.FromBase64String(encryptedValue);
            return Encoding.UTF8.GetString(data);
        }
        catch
        {
            // Falls der Wert nicht entschlüsselt werden kann, wird er unverändert zurückgegeben
            return encryptedValue;
        }
    }
}    