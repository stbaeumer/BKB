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

Global.H1 = "BKB.exe | https://github.com/stbaeumer/BKB | GPLv3 | 18.03.2025";
Global.User = Environment.UserName;

Global.CodeSpace = Global.RunningInCodeSpace();
Global.DefaultBackground = Console.BackgroundColor;
Global.PfadExportdateien = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

var bkbJsonPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BKB.json");

do
{
    Global.DisplayHeader(Global.Header.H1, Global.H1, Global.Protokollieren.Nein);    
        
    if (!File.Exists(bkbJsonPath))
{     
    var bkbJsonContent = new
    {
        Global.PfadExportdateien,
        Global.PfadSchilddateien,
        Kalenderfilter = Global.EncryptValue(""),    
        Auswahl = Global.EncryptValue("1"),
        Klassen = Global.EncryptValue("HBG"),
        Vergleich = Global.EncryptValue("n"),
        Kennwort = Global.EncryptValue(""),
        InputFolder = Global.EncryptValue(""),
        OutputFolder = Global.EncryptValue(""),
        Halbjahreszeugnisdatum = Global.EncryptValue(DateTime.Now.ToString("dd.MM.yyyy")),
        Halbjahreskonferenzdatum = Global.EncryptValue(DateTime.Now.ToString("dd.MM.yyyy")),
        Jahreszeugnisdatum = Global.EncryptValue(DateTime.Now.ToString("dd.MM.yyyy")),
        Jahreskonferenzdatum = Global.EncryptValue(DateTime.Now.ToString("dd.MM.yyyy")),
        Sprechtagsdatum = Global.EncryptValue(DateTime.Now.ToString("dd.MM.yyyy")),
        WikiUrl = Global.EncryptValue("http://192.168.134.10/lib/exe/xmlrpc.php"),
        WikiJsonUser = Global.EncryptValue("root"),
        WikiJsonUserKennwort = Global.EncryptValue(""),
        WikiSprechtagKleineAenderung = Global.EncryptValue(""),
        Zaehlerfach = Global.EncryptValue("j"),
        MariaUser = Global.EncryptValue(""),
        MariaIp = Global.EncryptValue(""),
        MariaPort = Global.EncryptValue(""),
        MariaDb = Global.EncryptValue(""),
        MariaPw = Global.EncryptValue(""),
        FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt = Global.EncryptValue(""),
        MaximaleAnzahlFehlstundenProTag = Global.EncryptValue("8"),
        Abschnitt = Global.EncryptValue("1"),
        Chat = Global.EncryptValue(""),
        AusbuchenNachWievielTagen = Global.EncryptValue(""),
        DatenimportLetztesDatum = Global.EncryptValue(DateTime.Now.ToString("dd.MM.yyyy")),
        MaxDateiAlter = Global.EncryptValue("6"),
        AktSj = Global.EncryptValue(""),
        Klasse = Global.EncryptValue(""),
        MailDomain = Global.EncryptValue("@students.berufskolleg-borken.de"),
        ZipKennwort = Global.EncryptValue(""),
        ConnectionStringUntis = Global.EncryptValue(""),
        EinstellungenVorgenommen = Global.EncryptValue("n"),
        SmtpServer = Global.EncryptValue("smtp.office365.com"),
        SmtpUser = Global.EncryptValue("stefan.baeumer@berufskolleg-borken.de"),
        SmtpPassword = Global.EncryptValue(""),
        SmtpPort = Global.EncryptValue("587"),
        NetmanMailReceiver = Global.EncryptValue("stefan.baeumer@berufskolleg-borken.de")
    };
    var json = JsonSerializer.Serialize(bkbJsonContent, new JsonSerializerOptions { WriteIndented = true });        
    File.WriteAllText(bkbJsonPath, json);                        
    Global.EinstellungenDurchlaufen(new ConfigurationBuilder().SetBasePath(Path.GetDirectoryName(bkbJsonPath)).AddJsonFile(Path.GetFileName(bkbJsonPath), optional: false, reloadOnChange: true).Build());
}

    var configuration = new ConfigurationBuilder()
    .SetBasePath(Path.GetDirectoryName(bkbJsonPath))
    .AddJsonFile(Path.GetFileName(bkbJsonPath), optional: false, reloadOnChange: true)
    .Build();

// Dynamisch alle Werte aus der Konfiguration den globalen Variablen zuweisen
foreach (var property in typeof(Global).GetProperties())
{
    var configValue = configuration[property.Name];
    if (configValue != null)
    {
        // Entschlüsseln des Wertes
        var decryptedValue = Global.DecryptValue(configValue);

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
    menue.AuswahlKonsole(configuration);
    Global.WeiterMitAnykey(configuration);

} while (true);