using System.Globalization;
using Microsoft.Data.SqlClient;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

public class Lehrers : List<Lehrer>
{
    private Dateien dateien;

    public Lehrers()
    {
    }
    public Lehrers(Dateien dateien)
    {
        List<dynamic>? lul = dateien.GetMatchingList("lehrkraefte", null, null);
        
        var liste = new List<dynamic>();
        
        foreach (var rec in lul)
        {
            var dict = (IDictionary<string, object>)rec;

            var l = new Lehrer();
            l.Kürzel = dict["InternKrz"].ToString();
            l.Nachname = dict["Nachname"].ToString();
            l.Vorname = dict["Vorname"].ToString();
            l.Mail = dict["E-Mail"].ToString();
            l.Titel = dict["Titel"].ToString();
            this.Add(l);
        }
    }

    public Lehrers(int periode, Raums raums)
    {
        var keinGeburtsdatum = new List<string?>();
        var keinGeschlecht = new List<string?>();

        using (var odbcConnection = new SqlConnection(Global.ConnectionStringUntis))
        {
            try
            {
                string queryString = @"SELECT DISTINCT 
Teacher.Teacher_ID, 
Teacher.Name,
Teacher.Longname, 
Teacher.FirstName,
Teacher.Email,
Teacher.ROOM_ID,
Teacher.Title,
Teacher.PlannedWeek,
Teacher.Flags,
Teacher.BirthDate,
Teacher.Text2
 FROM Teacher
WHERE (((SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] + ") AND  ((TERM_ID)=" + periode +
                                     ") AND ((Teacher.SCHOOL_ID)=177659) AND (((Teacher.Deleted)='false'))) ORDER BY Teacher.Name;";

                SqlCommand odbcCommand = new SqlCommand(queryString, odbcConnection);
                odbcConnection.Open();
                SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                while (sqlDataReader.Read())
                {
                    var lehrer = new Lehrer();

                    lehrer.IdUntis = sqlDataReader.GetInt32(0);
                    lehrer.Kürzel = Global.SafeGetString(sqlDataReader, 1);
                    lehrer.Nachname = Global.SafeGetString(sqlDataReader, 2);

                    if (lehrer.Nachname != "")
                    {
                        try
                        {
                            lehrer.Flags = Global.SafeGetString(sqlDataReader, 10);
                            lehrer.Vorname = Global.SafeGetString(sqlDataReader, 3);
                            lehrer.Mail = Global.SafeGetString(sqlDataReader, 4);
                            lehrer.Raum =
                                (from r in raums where r.IdUntis == sqlDataReader.GetInt32(5) select r.Raumnummer)
                                .FirstOrDefault();
                            lehrer.Titel = Global.SafeGetString(sqlDataReader, 6);
                            lehrer.Text2 = Global.SafeGetString(sqlDataReader, 10);
                            lehrer.Deputat = Convert.ToDouble(sqlDataReader.GetInt32(7)) / 1000;
                            lehrer.Geschlecht = Global.SafeGetString(sqlDataReader, 8).Contains("W") ? "w" : "m";

                            if (lehrer.Geschlecht != "w" && lehrer.Geschlecht != "m")
                            {
                                keinGeschlecht.Add(lehrer.Kürzel);
                            }

                            try
                            {
                                lehrer.Geburtsdatum = DateTime.ParseExact(sqlDataReader.GetInt32(9).ToString(),
                                    "yyyyMMdd", CultureInfo.InvariantCulture);
                            }
                            catch (Exception)
                            {
                                // Bei Nicht-Lehrern ist das Geb.Dat. egal
                                if (lehrer.Deputat > 0)
                                {
                                    if (lehrer.Kürzel != "MOR" && lehrer.Kürzel != "TIS")
                                    {
                                        keinGeburtsdatum.Add(lehrer.Kürzel);
                                    }
                                }
                            }

                            if (lehrer.Geburtsdatum.Year > 1)
                            {
                                lehrer.AlterAmErstenSchultagDiesesJahres =
                                    lehrer.GetAlterAmErstenSchultagDiesesJahres();
                                lehrer.ProzentStelle = lehrer.GetProzentStelle();
                            }

                            this.Add(lehrer);
                        }
                        catch (Exception ex)
                        {
                            Global.ZeileSchreiben(lehrer.Nachname, this.Count().ToString(),
                                ConsoleColor.Red,ConsoleColor.Gray);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception(ex.ToString());
            }
            finally
            {
                odbcConnection.Close();
                Global.ZeileSchreiben("Lehrer", this.Count().ToString(),
                    ConsoleColor.Green, ConsoleColor.White);
                foreach (var kuerzel in keinGeburtsdatum)
                {
                    //Global.ZeileSchreiben(1, (kuerzel + " ").PadRight(70, '.'), "kein Geburtsdatum", null);
                }

                foreach (var kuerzel in keinGeschlecht)
                {
                    //Global.ZeileSchreiben(1, (kuerzel + " ").PadRight(70, '.'), "kein Geschlecht", null);
                }
            }
        }
    }

    public void CheckAltersermäßigung(Common.Anrechnungen anrechnungen)
    {
        bool diesesJahr = false;

        Console.WriteLine("Für welches Schuljahr soll die Altersermäßigung errechnet werden?");
        Console.WriteLine("1: " + Global.AktSj[0] + "/" + Global.AktSj[1]);
        Console.WriteLine("2: " + Global.AktSj[1] + "/" + (Convert.ToInt32(Global.AktSj[1]) + 1));

        Console.Write("Geben Sie 1 oder 2 ein: ");
        var x = Console.ReadLine();

        if (x == "2")
            diesesJahr = true;

        foreach (var lehrer in this)
        {
            lehrer.AusgeschütteteAltersermäßigung = anrechnungen
                .Where(x => x.Grund == 200 && x.Lehrer.Kürzel == lehrer.Kürzel).Select(x => x.Wert).FirstOrDefault();

            lehrer.CheckAltersermäßigung(diesesJahr);
        }
    }

    public void GetTeamsUrl(IEnumerable<dynamic> verschiedeneLehrerKuerzel, string klassen)
    {
        var url = "https://teams.microsoft.com/l/chat/0/0?users=";

        if (Global.User != "HS")
        {
            url += "Kerstin.hues@berufskolleg-borken.de,";
        }

        if (Global.User != "LS")
        {
            url += "klaus.lienenklaus@berufskolleg-borken.de,";
        }

        if (Global.User != "BM")
        {
            url += "stefan.baeumer@berufskolleg-borken.de,";
        }

        if (!Global.User.ToUpper().StartsWith("MOR"))
        {
            url += "ursula.moritz@berufskolleg-borken.de,";
        }

        int anzahlTeamsChat = 0;

        // Der Teams-Chat ist auf Zeichen begrenzt.
        List<string> überzählige = new List<string>();

        foreach (var kuerzel in verschiedeneLehrerKuerzel)
        {
            var mail = (from l in this where l.Kürzel == kuerzel select l.Mail).FirstOrDefault();

            if (!string.IsNullOrEmpty(mail) && !url.Contains(mail))
            {
                // Es können maximal 7 in den Chat URL übergeben werden.
                anzahlTeamsChat++;

                if (anzahlTeamsChat < 7)
                {
                    url += mail + ",";
                }
                else
                {
                    if (!überzählige.Contains(mail))
                    {
                        überzählige.Add(mail);
                    }
                }
            }
        }

        Console.WriteLine(" ");
        Console.WriteLine("  Link zum Teams-Chat mit den LuL der Klasse " + klassen + ":");
        Console.WriteLine(" " + url.TrimEnd(','));

        if (überzählige.Count > 0)
        {
            Console.WriteLine(
                "  A C H T U N G: Folgende LuL müssen dem Teams-Chat zusätzlich manuell hinzugefügt werden:");
            foreach (var item in überzählige)
            {
                Console.WriteLine("   " + item);
            }
        }
    }

    public void OffeneKlassenbuchEinträgeMahnen(string dateiName)
    {
        List<string> lehrer = new List<string>();

        using (var pdfDocument = PdfDocument.Open(dateiName))
        {
            foreach (Page page in pdfDocument.GetPages())
            {
                foreach (var word in page.GetWords())
                {
                    // Prüfe, ob die linke x-Koordinate (BoundingBox.Left) ungefähr 100 ist
                    if (Math.Abs(word.BoundingBox.Left - 100) < 0.1)
                    {
                        if (word.Text != "Lehrkraft"){
                            lehrer.Add(word.Text);
                        }                        
                    }
                }                
            }
        }

        // Gib die 3 häufigsten Nennungen aus der Liste "lehrer" aus
        var topLehrer = lehrer.GroupBy(x => x)
                      .OrderByDescending(g => g.Count())
                      .Take(10)
                      .Select(g => new { Name = g.Key, Count = g.Count() });

        Global.ZeileSchreiben("TOP10 Offene Klassenbuch-Einträge", "Häufigkeit", ConsoleColor.Black, ConsoleColor.Blue);
                
        foreach (var item in topLehrer)
        {
            Global.ZeileSchreiben($"{item.Name}", $"{item.Count}", ConsoleColor.Blue, ConsoleColor.Black);        
        }   

        
        Console.WriteLine("  Jetzt per Mail senden? [J/n]");
        var x = Console.ReadKey().Key;
        if (x == ConsoleKey.J || x == ConsoleKey.Enter)
        {
            int i = 1;
            foreach (var item in topLehrer)
            {
                var le = (from l in this where l.Kürzel == item.Name select l).FirstOrDefault();
                
                if(le != null)
                {
                    var body = "Guten Morgen " + le.Titel+ le.Vorname + " " + le.Nachname + ",\n\n";
                    body += "es liegen " + item.Count + " offene Klassenbuch-Einträge vor, die Ihrer Verantwortung zugeordnet sind. ";
                    body += "Bitte kümmern Sie sich zeitnah um die Bearbeitung dieser Einträge.\n\n";
                    body += "Vielen Dank für Ihre Unterstützung.\n\n";
                    body += "Mit freundlichen Grüßen\n\n";
                    body += "Ihr Webuntis-Team";
                    
                    var mail = new Mail();                        
                    mail.Senden($"😐 Platz #{i} in der Liste der offenen Klassenbuch-Einträge", 
                "webuntis@berufskolleg-borken.de", 
                body, 
                null, 
                le.Mail); 
                    Global.ZeileSchreiben($"Name: {item.Name}", $"gesendet", ConsoleColor.Green, ConsoleColor.White);        
                }
                
                i++;
            }
        }else
        {
            Console.WriteLine("  Sie haben sich gegen den Mailversand entschieden.");
        }
    }
}