using System.Runtime.InteropServices;
using System.Text.Json;
using Common;
using Microsoft.Extensions.Configuration;

try
{
    Global.PfadSchilddateien = @"\\fs01\Schild-NRW\Datenaustausch\";
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

Global.H1 = "BKB.exe | https://github.com/stbaeumer/BKB | GPLv3 | 17.03.2025";
Global.User = Environment.UserName;

Global.CodeSpace = Global.RunningInCodeSpace();
Global.DefaultBackground = Console.BackgroundColor;
Global.PfadExportdateien = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

var bkbJsonPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BKB.json");

do
{
    Global.DisplayHeader(Global.Header.H1, Global.H1, Global.Protokollieren.Nein);    
    var configuration = new ConfigurationBuilder().SetBasePath(Path.GetDirectoryName(bkbJsonPath)).AddJsonFile(Path.GetFileName(bkbJsonPath), optional: false, reloadOnChange: true).Build();       
    
    if(!File.Exists(bkbJsonPath))
    {     
        var bkbJsonContent = new
        {
            Global.PfadExportdateien,
            Global.PfadSchilddateien,
            Kalenderfilter = "",    
            Auswahl = "1",
            Klassen = "HBG",
            Vergleich = "n",
            Kennwort = "1234",
            InputFolder = "",
            OutputFolder = "",
            Halbjahreszeugnisdatum = DateTime.Now.ToString("dd.MM.yyyy"),
            Halbjahreskonferenzdatum = DateTime.Now.ToString("dd.MM.yyyy"),
            Jahreszeugnisdatum = DateTime.Now.ToString("dd.MM.yyyy"),
            Jahreskonferenzdatum = DateTime.Now.ToString("dd.MM.yyyy"),
            Sprechtagsdatum = DateTime.Now.ToString("dd.MM.yyyy"),
            WikiUrl = "http://192.168.134.10/lib/exe/xmlrpc.php",
            WikiJsonUser = "root",
            WikiJsonUserKennwort = "",
            WikiSprechtagKleineÄnderung = "",
            Zaehlerfach = "j",
            MariaUser = "",
            MariaIp = "",
            MariaPort = "",
            MariaDb = "",
            MariaPw = "",
            FehlzeitenWährendDerLetztenTagBleibenUnberücksichtigt = "",
            MaximaleAnzahlFehlstundenProTag = "8",
            Abschnitt = "1",
            Chat = "",
            AusbuchenNachWievielTagen = "",
            DatenimportLetztesDatum = DateTime.Now.ToString("dd.MM.yyyy"),
            MaxDateiAlter = "6",
            AktSj = "",
            Klasse = "",
            MailDomain = "@students.berufskolleg-borken.de",
            ZipKennwort = "",
            ConnectionStringUntis = "",
            EinstellungenVorgenommen = "n",
            SmtpServer = "smtp.office365.com",
            SmtpUser = "stefan.baeumer@berufskolleg-borken.de",
            SmtpPassword = "",
            SmtpPort = "587",
            NetmanMailReceiver = "stefan.baeumer@berufskolleg-borken.de"
        };
        var json = JsonSerializer.Serialize(bkbJsonContent, new JsonSerializerOptions { WriteIndented = true });        
        File.WriteAllText(bkbJsonPath, json);                
        Global.EinstellungenDurchlaufen(configuration);
    }

    // Dynamisch alle Werte aus der Konfiguration den globalen Variablen zuweisen
    foreach (var property in typeof(Global).GetProperties())
    {
        var configValue = configuration[property.Name];
        if (configValue != null)
        {
            var propertyType = property.PropertyType;
            if (propertyType == typeof(int))
            {
                if (configValue.ToString().Length == 0)
                {
                    configValue = "0";
                }
                property.SetValue(null, Convert.ToInt32(configValue));
            }
            else if (propertyType == typeof(bool))
            {
                property.SetValue(null, Convert.ToBoolean(configValue));
            }
            else if (propertyType == typeof(DateTime))
            {
                if(configValue.ToString().Length == 0)
                {
                    configValue = DateTime.Now.ToString("dd.MM.yyyy");
                }
                property.SetValue(null, Convert.ToDateTime(configValue));
            }
            else if (propertyType == typeof(string))
            {
                property.SetValue(null, configValue);
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
    menue.AuswahlKonsole(configuration);
    Global.WeiterMitAnykey(configuration);

} while (true);