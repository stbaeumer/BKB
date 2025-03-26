using System.Runtime.InteropServices;
using System.Text.Json;
using Common;
using Microsoft.Extensions.Configuration;

try
{
    Global.PfadSchilddateien = @"\\fs01\SchILD-NRW\Datenaustausch\";
    Console.Clear();

    // Nur für Windows:
    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
    {
        Console.WindowHeight = 130;
        Console.WindowWidth = 130;

        // Ensure buffer is large enough
        if (Console.BufferWidth < Console.WindowWidth)
            Console.BufferWidth = Console.WindowWidth;

        if (Console.BufferHeight < Console.WindowHeight)
            Console.BufferHeight = Console.WindowHeight;

        // Now safely set window size
        Console.SetWindowSize(Console.WindowWidth, Console.WindowHeight);        
    }
    else
    {
        Global.PfadSchilddateien = @"/home/stefan/Windows/Datenaustausch";
    }
}
catch (ArgumentOutOfRangeException)
{
}

// Setze den Hintergrund auf weiß und den Vordergrund auf schwarz
Global.DefaultBackground = ConsoleColor.White;
Console.ForegroundColor = ConsoleColor.Black;
Console.Clear();

Global.H1 = "BKB-Tool | https://github.com/stbaeumer/BKB | GPLv3 | 26.03.2025";
Global.User = Environment.UserName;

Global.CodeSpace = Global.RunningInCodeSpace();
Global.DefaultBackground = Console.BackgroundColor;
Global.PfadExportdateien = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

var bkbJsonPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BKB.json");

do
{
    Global.DisplayHeader(Global.Header.H1, Global.H1, Global.Protokollieren.Nein);
        
    JsonDateiErstellenOderUpdaten(bkbJsonPath);

    var configuration = GetConfiguration(bkbJsonPath);

    DynamischAlleWerteAusDerKonfigurationDenGlobalenVariablenZuweisen(configuration);
            
    Global.ExportAusSchildVerschieben(configuration);

    var raums = new Raums();
    var lehrers = new Lehrers();
    var klassen = new Klassen();
    var anrechnungen = new Anrechnungen();

    try
    {
        if (!string.IsNullOrEmpty(Global.ConnectionStringUntis) && Global.ConnectionStringUntis.Length > 1)
        {
            Global.ZeileSchreiben("Daten aus Untis:", "mit DB verbunden", ConsoleColor.White,ConsoleColor.Green);
        }
        var periodes = new Periodes();
        if (periodes.Count > 0)
        {
            var periode = periodes.GetAktuellePeriode();
            //var feriens = new Feriens();
            raums = new Raums(periode);
            lehrers = new Lehrers(periode, raums);
            klassen = new Klassen(periode, lehrers, raums);
            anrechnungen = new Anrechnungen(lehrers);    
        }
    }
    catch 
    {
        Console.WriteLine("Keine Verbindung zu Untis.");
    }

    var dateien = new Dateien();
    dateien.GetInteressierendeDateienMitAllenEigenschaften();    
    Global.ZeileSchreiben("Dateien einlesen", Global.PfadExportdateien, ConsoleColor.Black, ConsoleColor.Magenta);
    dateien.GetZeilen();        
    var menue = MenueHelper.Einlesen(dateien, klassen, lehrers, configuration, anrechnungen, raums);
    if (menue == null) continue;
    var menueintrag = menue.AuswahlKonsole(configuration);
    Global.WeiterMitAnykey(configuration, menueintrag);
} while (true);

void DynamischAlleWerteAusDerKonfigurationDenGlobalenVariablenZuweisen(IConfiguration configuration)
{
    foreach (var property in typeof(Global).GetProperties())
    {
        var configValue = configuration[property.Name];
        if (configValue != null)
        {
            // Entschlüsseln des Wertes
            var decryptedValue = Global.Entschluesseln(configValue);

            var propertyType = property.PropertyType;
            if (propertyType == typeof(int))
            {
                if (decryptedValue.Length == 0)
                {
                    decryptedValue = "0";
                }
                property.SetValue(null, Convert.ToInt32(decryptedValue));
            }
            else if (propertyType == typeof(bool))
            {
                property.SetValue(null, Convert.ToBoolean(decryptedValue));
            }
            else if (propertyType == typeof(DateTime))
            {
                if (decryptedValue.Length == 0)
                {
                    decryptedValue = DateTime.Now.ToString("dd.MM.yyyy");
                }
                property.SetValue(null, Convert.ToDateTime(decryptedValue));
            }
            else if (propertyType == typeof(string))
            {
                property.SetValue(null, decryptedValue);
            }
            // Fügen Sie hier weitere Typen hinzu, falls erforderlich
        }
    }
}

IConfiguration GetConfiguration(string bkbJsonPath)
{
    return new ConfigurationBuilder()
    .SetBasePath(Path.GetDirectoryName(bkbJsonPath))
    .AddJsonFile(Path.GetFileName(bkbJsonPath), optional: false, reloadOnChange: true)
    .Build();
}

void JsonDateiErstellenOderUpdaten(string bkbJsonPath)
{
    if (!File.Exists(bkbJsonPath))
    {
        // Datei erstellen, wenn sie nicht existiert
        var bkbJsonContent = CreateBkbJsonContent();
        var json = JsonSerializer.Serialize(bkbJsonContent, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(bkbJsonPath, json);
        Global.EinstellungenDurchlaufen(new ConfigurationBuilder().SetBasePath(Path.GetDirectoryName(bkbJsonPath)).AddJsonFile(Path.GetFileName(bkbJsonPath), optional: false, reloadOnChange: true).Build());
    }
    else
    {
        // Datei existiert, prüfen und fehlende Keys ergänzen
        var existingContent = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(bkbJsonPath));
        // Standard-Keys aus CreateBkbJsonContent
        var defaultContent = JsonSerializer.SerializeToElement(CreateBkbJsonContent()).EnumerateObject();
        var defaultKeys = defaultContent.Select(property => property.Name).ToHashSet();

        // Entferne nicht mehr benötigte Keys
        var keysToRemove = existingContent.Keys.Except(defaultKeys).ToList();
        foreach (var key in keysToRemove)
        {
            existingContent.Remove(key);
        }

        // Ergänze fehlende Keys
        foreach (var property in defaultContent)
        {
            if (!existingContent.ContainsKey(property.Name))
            {
                existingContent[property.Name] = property.Value.GetString();
            }
        }

        // Datei mit aktualisierten Inhalten speichern
        var updatedJson = JsonSerializer.Serialize(existingContent, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(bkbJsonPath, updatedJson);
    }
}

static object CreateBkbJsonContent()
{
    return new
    {
        ConnectionStringUntis = Global.Verschluesseln(@""),
        SchipsPasswort = Global.Verschluesseln(""),
        ZeugnisPasswort = Global.Verschluesseln(""),
        ZeugnisUrl = Global.Verschluesseln("https://bkb.wiki/notenlisten:start"),
        SchipsUrl = Global.Verschluesseln("https://bkb.wiki/statistik:schips:start"),
        SmtpPassword = Global.Verschluesseln(""),
        PdfKennwort = Global.Verschluesseln("Geheim123"),
        SchipsOderZeugnisseOderAnderePdfs = Global.Verschluesseln("1"),
        Global.PfadExportdateien,
        Global.PfadSchilddateien,
        Kalenderfilter = Global.Verschluesseln(""),
        Auswahl = 1,
        Klassen = Global.Verschluesseln("HBG"),
        Vergleich = Global.Verschluesseln("n"),
        Kennwort = Global.Verschluesseln(""),        
        InputFolder = Global.Verschluesseln(""),
        OutputFolder = Global.Verschluesseln(""),
        Halbjahreszeugnisdatum = Global.Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
        Halbjahreskonferenzdatum = Global.Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
        Jahreszeugnisdatum = Global.Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
        Jahreskonferenzdatum = Global.Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
        Sprechtagsdatum = Global.Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
        WikiUrl = Global.Verschluesseln("http://192.168.134.10/lib/exe/xmlrpc.php"),
        WikiJsonUser = Global.Verschluesseln("root"),
        WikiJsonUserKennwort = Global.Verschluesseln(""),
        WikiSprechtagKleineAenderung = Global.Verschluesseln(""),
        Zaehlerfach = Global.Verschluesseln("j"),
        MariaUser = Global.Verschluesseln(""),
        MariaIp = Global.Verschluesseln(""),
        MariaPort = Global.Verschluesseln(""),
        MariaDb = Global.Verschluesseln(""),
        MariaPw = Global.Verschluesseln(""),
        FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt = Global.Verschluesseln(""),
        MaximaleAnzahlFehlstundenProTag = Global.Verschluesseln("8"),
        Abschnitt = Global.Verschluesseln("1"),
        Chat = Global.Verschluesseln(""),
        AusbuchenNachWievielTagen = Global.Verschluesseln(""),
        DatenimportLetztesDatum = Global.Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
        MaxDateiAlter = Global.Verschluesseln("6"),
        AktSj = Global.Verschluesseln(""),
        Klasse = Global.Verschluesseln(""),
        MailDomain = Global.Verschluesseln("@students.berufskolleg-borken.de"),
        ZipKennwort = Global.Verschluesseln("Geheim123"),        
        EinstellungenVorgenommen = Global.Verschluesseln("n"),
        SmtpServer = Global.Verschluesseln("smtp.office365.com"),
        SmtpUser = Global.Verschluesseln("webuntis@berufskolleg-borken.de"),
        SmtpPort = Global.Verschluesseln("587"),
        NetmanMailReceiver = Global.Verschluesseln("stefan.baeumer@berufskolleg-borken.de"),        
        Betreff = Global.Verschluesseln("Betreff"),
        Body = Global.Verschluesseln("Guten Morgen [Lehrer],\n\nbitte beachten Sie den Anhang.\n\nErläuterungen dazu finden Sie hier: https://bkb.wiki/konzepte:stundenplanungskonzept#information_aller_lehrkraefte_per_mail \n\nViele Grüße aus der Schulverwaltung"),
    };
}