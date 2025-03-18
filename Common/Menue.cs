using System.Diagnostics;
using Common;
using Microsoft.Extensions.Configuration;

public class Menue : List<Menüeintrag>
{
    public int AusgewaehlterMenueEintrag { get; set; }
    private Menüeintrag AusgewaehlterEintrag { get; set; }
    public Dateien Quelldateien { get; set; }
    public Klassen Klassen { get; set; }
    public Lehrers Lehrers { get; set; }
    
    /// <summary>
    /// Menü aufbauen und Dateien Filtern
    /// </summary>
    public Menue(Dateien quelldateien, Klassen klassen, Lehrers lehrers, Students students, List<Menüeintrag> menueEintraege)
    {
        Quelldateien = quelldateien;
        Klassen = klassen;
        Lehrers = lehrers;
        
        if (!string.IsNullOrEmpty(Global.AktuellerPfad))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = Global.AktuellerPfad,
                UseShellExecute = true,
                Verb = "open"
            });
            Global.AktuellerPfad = "";
        }

        //Global.WeiterMitAnyKey();
        //Global.DisplayHeader(Global.Header.H1, Global.H1!, Global.Protokollieren.Nein);


        Global.Protokoll = new List<string>();
        AddRange(menueEintraege);
        //AddRange(menueEintraege.Where(x => !x.MenüeintragAusblenden));
    }

    public void AuswahlKonsole(IConfiguration configuration)
    {
        Global.ZeileSchreiben("Auswahl treffen:", "x: Einstellungen ändern; y: Dokumentation online öffnen", ConsoleColor.Black, ConsoleColor.Yellow);
        
        var documentsFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        var configPath = Path.Combine(documentsFolderPath, "BKB.json");

        Console.ForegroundColor = ConsoleColor.Yellow;
        //Console.WriteLine("".PadRight(Console.WindowWidth, '='));

        var xx = IndexOf(this.FirstOrDefault(menueeintrag => menueeintrag.Titel == configuration["Auswahl"])!);

        AusgewaehlterMenueEintrag = xx >= 0 ? Math.Max(1, xx + 1) : 1;

        bool fehlermeldung = false;

        for (int i = 0; i < this.Count(); i++)
        {
            Console.ResetColor();
            if (this[i].Quelldateien.Count > 0 && this[i].Quelldateien.Any(q => !string.IsNullOrEmpty(q.Fehlermeldung)))
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine(" " + (i + 1).ToString().PadLeft(3) + ". " + (this[i].Titel + " *)").PadRight(13));   
                fehlermeldung = true;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(" " + (i + 1).ToString().PadLeft(3) + ". " + this[i].Titel.PadRight(13));   
            }                     
        }

        if(fehlermeldung)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("");
            Console.WriteLine("     *) Der Menüeintrag kann nicht gewählt werden. Dateien / Tabellen fehlen oder sind leer.");
            Console.WriteLine("");
        }

        var wiederhole = true;
        do
        {
            Console.WriteLine("");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("    Ihre Auswahl [" + AusgewaehlterMenueEintrag + "] : ");
            Console.ResetColor();
            var eingabe = Console.ReadLine();

            if (eingabe != "ö")
            {
                if (eingabe != "x")
                {
                    if (eingabe != "y")
                    {
                        if (eingabe == "" && AusgewaehlterMenueEintrag.ToString() != "")
                        {
                            eingabe = AusgewaehlterMenueEintrag.ToString();
                        }

                        var nummer = 0;

                        if (int.TryParse(eingabe, out nummer))
                        {
                            // Überprüfen, ob die Zahl im gültigen Bereich liegt
                            if (nummer >= 1 && nummer <= this.Count)
                            {
                                this.AusgewaehlterMenueEintrag = nummer;
                                var x = this[AusgewaehlterMenueEintrag - 1].Titel;

                                if(this[AusgewaehlterMenueEintrag - 1].Quelldateien.Any(q => !string.IsNullOrEmpty(q.Fehlermeldung)))
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Global.ZeileSchreiben("Der Menüeintrag kann nicht gewählt werden:", "Hinweise beachten:", ConsoleColor.Black, ConsoleColor.Red);
                                    Console.WriteLine("");
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    
                                    
                                    foreach (var datei in this[AusgewaehlterMenueEintrag - 1].Quelldateien.Where(q => !string.IsNullOrEmpty(q.Fehlermeldung)))
                                    {
                                        Global.ZeileSchreiben(datei.Dateiname, datei.Fehlermeldung, ConsoleColor.Black, ConsoleColor.Red);
                                        Global.DisplayCenteredBox(datei.Hinweise.ToList(), 97);
                                    }
                                    
                                    
                                    Console.ResetColor();
                                    wiederhole = true;
                                    continue;
                                }

                                Global.Speichern("Auswahl", x);
                                wiederhole = false;
                            }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("     Die Zahl muss zwischen 1 und " + this.Count +
                                                  " liegen. Bitte versuchen Sie es erneut.");
                                Console.ResetColor();
                                wiederhole = true;
                                continue;
                            }
                        }
                        else
                        {
                            if (eingabe == "" && AusgewaehlterMenueEintrag >= 1 &&
                                AusgewaehlterMenueEintrag <= this.Count) continue;
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("     Die Zahl muss zwischen 1 und " + this.Count +
                                              " liegen. Bitte versuchen Sie es erneut.");
                            Console.ResetColor();
                        }   
                    }
                    else
                    {
                        Global.OpenWebseite("");
                        wiederhole = true;
                        continue;
                    }
                }
                else
                {
                    Global.EinstellungenDurchlaufen(configuration);
                    return;
                }
            }
            else
            {
                Global.OpenCurrentFolder();
                wiederhole = true;
                continue;
            }
        } while (wiederhole);

        Global.DisplayHeader(Global.Header.H1, Global.H1, Global.Protokollieren.Nein);
        Global.DisplayHeader(Global.Header.H2, this[AusgewaehlterMenueEintrag - 1].Titel, Global.Protokollieren.Nein);
        Global.DisplayCenteredBox(this[AusgewaehlterMenueEintrag - 1].Beschreibung, 97);
        this[AusgewaehlterMenueEintrag - 1].Ausführen();
    }
}