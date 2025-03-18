using System.Globalization;
using System.Text.RegularExpressions;
using Common;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

public partial class Student
{
    public int F2PlusF3;
    public string? Foto { get; set; } = string.Empty;
    public string Aktjahrgang { get; set; } = string.Empty;
    public string AktuellerAbschnitt { get; set; } = string.Empty;
    public string AktuellesHalbjahr { get; set; } = string.Empty;
    public string Aufnahmedatum { get; set; } = string.Empty;
    public string Adressart { get; set; } = string.Empty;
    public string Anrede { get; set; } = string.Empty;
    public string Asv { get; set; } = string.Empty;
    public string AufnehmendeSchuleName { get; set; } = string.Empty;
    public string AufnehmendeSchuleOrt { get; set; } = string.Empty;
    public string AufnehmendeSchulePlz { get; set; } = string.Empty;
    public string AufnehmendeSchuleSchulnr { get; set; } = string.Empty;
    public string AufnehmendeSchuleStraße { get; set; } = string.Empty;
    public string Aussiedler { get; set; } = string.Empty;
    public string Ausweisnummer { get; set; } = string.Empty;
    public string Ausbildungsort { get; set; } = string.Empty;
    public string Austritt { get; set; } = string.Empty;
    public string AbschlussartAnEigenerSchuleStatistikKürzel { get; set; } = string.Empty;
    public string Bemerkungen { get; set; } = string.Empty;
    public string BeginnDerBildungsganges { get; set; } = string.Empty;
    public string BerufsschulpflichtErfüllt { get; set; } = string.Empty;
    public string BesMerkmal { get; set; } = string.Empty;
    public string BleibtAnSchule { get; set; } = string.Empty;
    public string Briefanrede { get; set; } = string.Empty;
    public string Berufsabschluss { get; set; } = string.Empty;
    public string Berufswechsel { get; set; } = string.Empty;
    public string Betreuung { get; set; } = string.Empty;
    public string BetreuerAbteilung { get; set; } = string.Empty;
    public string BetreuerAnrede { get; set; } = string.Empty;
    public string BetreuerEmail { get; set; } = string.Empty;
    public string BetreuerName { get; set; } = string.Empty;
    public string BetreuerTelefon { get; set; } = string.Empty;
    public string BetreuerTitel { get; set; } = string.Empty;
    public string BetreuerVorname { get; set; } = string.Empty;
    public string Betreuungslehrer { get; set; } = string.Empty;
    public string BetreuungslehrerAnrede { get; set; } = string.Empty;
    public string Bezugsjahr { get; set; } = string.Empty;
    public string Bkazvo { get; set; } = string.Empty;
    public string Branche { get; set; } = string.Empty;
    public string Bundesland { get; set; } = string.Empty;
    public string Einschulungsart { get; set; } = string.Empty;
    public string ElternteilZugezogen { get; set; } = string.Empty;
    public string Entlassjahrgang { get; set; } = string.Empty;
    public string Erzieher1Anrede { get; set; } = string.Empty;
    public string Erzieher1Briefanrede { get; set; } = string.Empty;
    public string Erzieher1Nachname { get; set; } = string.Empty;
    public string Erzieher1Titel { get; set; } = string.Empty;
    public string Erzieher1Vorname { get; set; } = string.Empty;
    public string Erzieher2Anrede { get; set; } = string.Empty;
    public string Erzieher2Briefanrede { get; set; } = string.Empty;
    public string Erzieher2Nachname { get; set; } = string.Empty;
    public string Erzieher2Titel { get; set; } = string.Empty;
    public string Erzieher2Vorname { get; set; } = string.Empty;
    public string ErzieherArtKlartext { get; set; } = string.Empty;
    public string ErzieherEmail { get; set; } = string.Empty;
    public string ErzieherErhältAnschreiben { get; set; } = string.Empty;
    public string ErzieherOrt { get; set; } = string.Empty;
    public string ErzieherOrtsteil { get; set; } = string.Empty;
    public string ErzieherPostleitzahl { get; set; } = string.Empty;
    public string ErzieherStraße { get; set; } = string.Empty;
    public string ExterneIdNummer { get; set; } = string.Empty;
    public string Fachklasse { get; set; } = string.Empty;
    public string FachklasseBezeichnung { get; set; } = string.Empty;
    public string FaxNr { get; set; } = string.Empty;
    public string Förderschwerpunkt1 { get; set; } = string.Empty;
    public string Förderschwerpunkt2 { get; set; } = string.Empty;
    public string? Geburtsdatum { get; set; } = string.Empty;
    public string Geburtsland { get; set; } = string.Empty;
    public string GeburtslandMutter { get; set; } = string.Empty;
    public string GeburtslandVater { get; set; } = string.Empty;
    public string Geburtsname { get; set; } = string.Empty;
    public string Geburtsort { get; set; } = string.Empty;
    public string Geschlecht { get; set; } = string.Empty;
    public string Gliederung { get; set; } = string.Empty;
    public string GsEmpfehlung { get; set; } = string.Empty;
    public string Hausnummer { get; set; } = string.Empty;
    public string HöchsterAllgAbschluss { get; set; } = string.Empty;
    public string Internatsplatz { get; set; } = string.Empty;
    public string IdSchild { get; set; } = string.Empty;
    public string JahrZuzug { get; set; } = string.Empty;
    public string JahrEinschulung { get; set; } = string.Empty;
    public string JahrSchulwechsel { get; set; } = string.Empty;
    public string Jahrgang { get; set; } = string.Empty;
    public string JahrgangInterneBezeichnung { get; set; } = string.Empty;
    public string Jva { get; set; } = string.Empty;
    public string Klasse { get; set; } = string.Empty;
    public string Klassenart { get; set; } = string.Empty;
    public string Klassenlehrer { get; set; } = string.Empty;
    public string KlassenlehrerAmtsbezeichnung { get; set; } = string.Empty;
    public string KlassenlehrerAnrede { get; set; } = string.Empty;
    public string KlassenlehrerName { get; set; } = string.Empty;
    public string KlassenlehrerTitel { get; set; } = string.Empty;
    public string KlassenlehrerVorname { get; set; } = string.Empty;
    public string Kreis { get; set; } = string.Empty;
    public string Koopklasse { get; set; } = string.Empty;
    public string LetzteSchuleName { get; set; } = string.Empty;
    public string LetzteSchuleOrt { get; set; } = string.Empty;
    public string LetzteSchulePlz { get; set; } = string.Empty;
    public string LetzterBerufsbezAbschlussKürzel { get; set; } = string.Empty;
    public string LsSchulform { get; set; } = string.Empty;
    public string Lsschulnummer { get; set; } = string.Empty;
    public string LsGliederung { get; set; } = string.Empty;
    public string LsFachklasse { get; set; } = string.Empty;
    public string Lsklassenart { get; set; } = string.Empty;
    public string Lsreformpdg { get; set; } = string.Empty;
    public string LsSschulentl { get; set; } = string.Empty;
    public string LsJahrgang { get; set; } = string.Empty;
    public string LsQual { get; set; } = string.Empty;
    public string Lsversetz { get; set; } = string.Empty;
    public string MailPrivat { get; set; } = string.Empty;
    public string MailSchulisch { get; set; } = string.Empty;
    public string Massnahmetraeger { get; set; } = string.Empty;
    public string MigrationshintergrundVorhanden { get; set; } = string.Empty;
    public string Nachname { get; set; } = string.Empty;
    public string Orgform { get; set; } = string.Empty;
    public string Ortsteil { get; set; } = string.Empty;
    public string Ort { get; set; } = string.Empty;
    public string Postleitzahl { get; set; } = string.Empty;
    public string Produktname { get; set; } = string.Empty;
    public string Produktversion { get; set; } = string.Empty;
    public string Reformpdg { get; set; } = string.Empty;
    public string Religionsanmeldung { get; set; } = string.Empty;
    public string Religionsabmeldung { get; set; } = string.Empty;
    public string KonfessionKlartext { get; set; } = string.Empty;
    public string Labk { get; set; } = string.Empty;
    public string Schwerstbehinderung { get; set; } = string.Empty;
    public string SchülerLeJahrDa { get; set; } = string.Empty;
    public string Schulwechselform { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Straße { get; set; } = string.Empty;
    public string Schulbesuchsjahre { get; set; } = string.Empty;
    public string Schulform { get; set; } = string.Empty;
    public string Zugezogen { get; set; } = string.Empty;
    public string Zeugnis { get; set; } = string.Empty;
    public string Volljährig { get; set; } = string.Empty;
    public string VoFoerderschwerpunkt2 { get; set; } = string.Empty;
    public string VoReformpdg { get; set; } = string.Empty;
    public string VoSchwerstbehindert { get; set; } = string.Empty;
    public string VoFoerderschwerp { get; set; } = string.Empty;
    public string VoJahrgang { get; set; } = string.Empty;
    public string VoKlassenart { get; set; } = string.Empty;
    public string VoOrgform { get; set; } = string.Empty;
    public string VoGliederung { get; set; } = string.Empty;
    public string VoKlasse { get; set; } = string.Empty;
    public string? Vorname { get; set; } = string.Empty;
    public string VorjahrC05AktjahrC06 { get; set; } = string.Empty;
    public string Versetzung { get; set; } = string.Empty;
    public string Vertragsende { get; set; } = string.Empty;
    public string Vertragsbeginn { get; set; } = string.Empty;
    public string Verkehrssprache { get; set; } = string.Empty;
    public string VoFachklasse { get; set; } = string.Empty;
    public string VoraussAbschlussdatum { get; set; } = string.Empty;
    public string TelefonNummernBemerkung { get; set; } = string.Empty;
    public string Telefonnummer2 { get; set; } = string.Empty;
    public string Telefonnummer { get; set; } = string.Empty;
    public string StaatsangehörigkeitSchlüssel { get; set; } = string.Empty;
    public string StaatsangehörigkeitKlartextAdjektiv { get; set; } = string.Empty;
    public string StaatsangehörigkeitKlartext { get; set; } = string.Empty;
    public string Sportbefreiung { get; set; } = string.Empty;
    public string Schwerpunkt { get; set; } = string.Empty;
    public string SchulpflichtErfüllt { get; set; } = string.Empty;
    public string SchulNummer { get; set; } = string.Empty;
    public string Schuljahr { get; set; } = string.Empty;
    public string SchulformFürSimExport { get; set; } = string.Empty;
    public string Zeugnisdatum { get; set; } = string.Empty;
    public string Adressmerkmal { get; set; } = string.Empty;
    public List<dynamic> Abwesenheiten { get; set; } = new List<dynamic>();
    public string AlleMaßnahmenUndVorgänge { get; set; } = string.Empty;
    public List<dynamic> Massnahmen { get; set; } = new List<dynamic>();
    public dynamic JuengsteMassnahmeInDiesemSj { get; set; }
    public int F2 { get; set; }
    public int F3 { get; set; }
    public int F2M { get; set; }
    public int F2MplusF3 { get; set; }
    public string MaßnahmenAlsWikiLinkAufzählung { get; set; } = string.Empty;
    public PdfSeiten PdfSeiten { get; set; } = new PdfSeiten();

    public void CreateFolderPdfDateien()
    {
        Zielordner = Global.OutputFolder + "/" + Nachname.Substring(0, 1) + "/" +
                     $"{Nachname}_{Vorname}_{Geburtsdatum}";

        if (string.IsNullOrEmpty(Nachname)) return;
        if (!Directory.Exists(Zielordner))
        {
            Directory.CreateDirectory(Zielordner);
        }
    }

    private string Zielordner { get; set; }
    public string Entlassdatum { get; set; }
    public bool FotoVorhanden { get; set; }
    public bool FotoBinary { get; set; }
    public int IdSchildInt { get; set; }
    public string KlasseString { get; set; }
    public string? KlasseWebuntis { get; set; }

    public string GetFehlstd(
        List<dynamic>? absencesPerStudent,
        int fehlzeitenWährendDerLetztenTagBleibenUnberücksichtigt
    )
    {
        try
        {
            var fehlstd = 0;

            foreach (var record in absencesPerStudent)
            {
                var dict = (IDictionary<string, object>)record;

                if (string.IsNullOrEmpty(dict["Schüler*innen"].ToString()) ||
                    !dict["Schüler*innen"].ToString()!.Contains(Nachname!)) continue;
                if (!dict["Schüler*innen"].ToString()!.Contains(Vorname!)) continue;
                if (string.IsNullOrEmpty(dict["Klasse"].ToString()) ||
                    !dict["Klasse"].ToString()!.Contains(Klasse)) continue;
                if (string.IsNullOrEmpty(dict["Fehlstd."].ToString())) continue;
                if ((DateTime.ParseExact(dict["Datum"].ToString()!, "dd.MM.yy",
                        System.Globalization.CultureInfo.InvariantCulture))
                    .AddDays(fehlzeitenWährendDerLetztenTagBleibenUnberücksichtigt) >=
                    DateTime.Now) continue;
                // Pro Tag werden maximal 8 Fehlstunden gezählt. Das ist wichtig, weil Webuntis bei ganztägigen Veranstaltungen 24 Stunden zählen würde.
                var webuntisFehlst = int.Parse(dict["Fehlstd."].ToString()!);
                if (webuntisFehlst > 10 && webuntisFehlst > Global.MaximaleAnzahlFehlstundenProTag)
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Global.ZeileSchreiben( this.Nachname + ", " + this.Vorname + " (" + this.Klasse + ")",
                        webuntisFehlst + " Fehlstunden am " + dict["Datum"].ToString() + " werden auf " +
                        Global.MaximaleAnzahlFehlstundenProTag + " begrenzt.", ConsoleColor.Yellow,ConsoleColor.Gray);
                    Console.ResetColor();
                }

                int fehlstundenAnDiesemTag = Math.Min(Global.MaximaleAnzahlFehlstundenProTag, webuntisFehlst);
                fehlstd = fehlstd + fehlstundenAnDiesemTag;
            }


            return fehlstd == 0 ? "" : fehlstd.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Console.ReadKey();
            return null!;
        }
    }

    public string GetUnentFehlstd(
        List<dynamic>? absencesPerStudent,
        int fehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt)
    {
        try
        {
            var fehlstd = 0;
            foreach (var record in absencesPerStudent)
            {
                var dict = (IDictionary<string, object>)record;

                if (!dict.ContainsKey("Schüler*innen") || dict["Schüler*innen"] == null ||
                    string.IsNullOrEmpty(dict["Schüler*innen"].ToString()) ||
                    !dict["Schüler*innen"].ToString()!.Contains(Nachname!)) continue;
                if (!dict.ContainsKey("Schüler*innen") || dict["Schüler*innen"] == null ||
                    string.IsNullOrEmpty(dict["Schüler*innen"].ToString()) ||
                    !dict["Schüler*innen"].ToString()!.Contains(Vorname!)) continue;
                if (!dict.ContainsKey("Klasse") || dict["Klasse"] == null ||
                    string.IsNullOrEmpty(dict["Klasse"].ToString()) ||
                    !dict["Klasse"].ToString()!.Contains(Klasse)) continue;
                if (!dict.ContainsKey("Status") || dict["Status"] == null ||
                    string.IsNullOrEmpty(dict["Status"].ToString()) ||
                    !dict["Status"].ToString()!.Contains("nicht entsch.")) continue;
                if (!dict.ContainsKey("Fehlstd.") || dict["Fehlstd."] == null ||
                    string.IsNullOrEmpty(dict["Fehlstd."].ToString())) continue;
                if ((DateTime.ParseExact(dict["Datum"].ToString()!, "dd.MM.yy",
                        System.Globalization.CultureInfo.InvariantCulture))
                    .AddDays(fehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt) >= DateTime.Now) continue;

                // Pro Tag und wenn Webuntis mehr als 10 Stunden zählt, werden maximal 8 Fehlstunden gezählt. Das ist wichtig, weil Webuntis bei ganztägigen Veranstaltungen 24 Stunden zählen würde.
                var webuntisFehlst = int.Parse(dict["Fehlstd."].ToString()!);
                if (webuntisFehlst > 10 && webuntisFehlst > Global.MaximaleAnzahlFehlstundenProTag)
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Global.ZeileSchreiben(this.Nachname + ", " + this.Vorname + " (" + this.Klasse + ")",
                        webuntisFehlst + " unentsch. Fehlstunden am " + dict["Datum"].ToString() + " werden auf " +
                        Global.MaximaleAnzahlFehlstundenProTag + " begrenzt.", ConsoleColor.Blue, ConsoleColor.White);
                    Console.ResetColor();
                }

                var fehlstundenAnDiesemTag = Math.Min(Global.MaximaleAnzahlFehlstundenProTag, webuntisFehlst);
                fehlstd += fehlstundenAnDiesemTag;
            }

            return fehlstd == 0 ? "" : fehlstd.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Console.ReadKey();
            return null!;
        }
    }

    string? FBereinigen(string? textinput)
    {
        string? text = textinput;

        text = text.ToLower(); // Nur Kleinbuchstaben
        text = FUmlauteBehandeln(text); // Umlaute ersetzen


        text = Regex.Replace(text, "-", "_"); //  kein Minus-Zeichen
        text = Regex.Replace(text, ",", "_"); //  kein Komma            
        text = Regex.Replace(text, " ", "_"); //  kein Leerzeichen
        // Text = Regex.Replace(Text, @"[^\w]", string.Empty);   // nur Buchstaben

        text = Regex.Replace(text, "[^a-z]", string.Empty); // nur Buchstaben

        text = text.Substring(0, 1); // Auf maximal 6 Zeichen begrenzen
        return text;
    }

    string? FUmlauteBehandeln(string? textinput)
    {
        string? text = textinput;

        // deutsche Sonderzeichen
        text = Regex.Replace(text, "[æ|ä]", "ae");
        text = Regex.Replace(text, "[Æ|Ä]", "Ae");
        text = Regex.Replace(text, "[œ|ö]", "oe");
        text = Regex.Replace(text, "[Œ|Ö]", "Oe");
        text = Regex.Replace(text, "[ü]", "ue");
        text = Regex.Replace(text, "[Ü]", "Ue");
        text = Regex.Replace(text, "ß", "ss");

        // Sonderzeichen aus anderen Sprachen
        text = Regex.Replace(text, "[ã|à|â|á|å]", "a");
        text = Regex.Replace(text, "[Ã|À|Â|Á|Å]", "A");
        text = Regex.Replace(text, "[é|è|ê|ë]", "e");
        text = Regex.Replace(text, "[É|È|Ê|Ë]", "E");
        text = Regex.Replace(text, "[í|ì|î|ï]", "i");
        text = Regex.Replace(text, "[Í|Ì|Î|Ï]", "I");
        text = Regex.Replace(text, "[õ|ò|ó|ô]", "o");
        text = Regex.Replace(text, "[Õ|Ó|Ò|Ô]", "O");
        text = Regex.Replace(text, "[ù|ú|û|µ]", "u");
        text = Regex.Replace(text, "[Ú|Ù|Û]", "U");
        text = Regex.Replace(text, "[ý|ÿ]", "y");
        text = Regex.Replace(text, "[Ý]", "Y");
        text = Regex.Replace(text, "[ç|č]", "c");
        text = Regex.Replace(text, "[Ç|Č]", "C");
        text = Regex.Replace(text, "[Ð]", "D");
        text = Regex.Replace(text, "[ñ]", "n");
        text = Regex.Replace(text, "[Ñ]", "N");
        text = Regex.Replace(text, "[š]", "s");
        text = Regex.Replace(text, "[Š]", "S");

        return text;
    }

    //public void GetFehlstd(Datei absencesPerStudent, List<int> aktSj, int abschnitt)
    //{
    //    try
    //    {
    //        Fehlstd = (from a in absencesPerStudent.Zeilen
    //                   where a[Array.IndexOf(absencesPerStudent.Kopfzeile, "Schüler*innen")].Contains(Nachname)
    //                   where a[Array.IndexOf(absencesPerStudent.Kopfzeile, "Schüler*innen")].Contains(Vorname)
    //                   select Convert.ToInt32(a[Array.IndexOf(absencesPerStudent.Kopfzeile, "Fehlstd.")])).Sum();
    //    }
    //    catch (Exception ex)
    //    {
    //        Console.WriteLine(ex.ToString());
    //        Console.ReadKey();
    //    }
    //}

    //public void GetUnentFehlstd(Datei absencesPerStudent, List<int> aktSj, int abschnitt)
    //{
    //    try
    //    {
    //        UnentschFehlstd = (from a in absencesPerStudent.Zeilen
    //                           where a[Array.IndexOf(absencesPerStudent.Kopfzeile, "Schüler*innen")].Contains(Nachname)
    //                           where a[Array.IndexOf(absencesPerStudent.Kopfzeile, "Schüler*innen")].Contains(Vorname)
    //                           where a[Array.IndexOf(absencesPerStudent.Kopfzeile, "Status")].Contains("nicht entsch.")
    //                           select Convert.ToInt32(a[Array.IndexOf(absencesPerStudent.Kopfzeile, "Fehlstd.")])).Sum();
    //    }
    //    catch (Exception ex)
    //    {
    //        Console.WriteLine(ex.ToString());
    //        Console.ReadKey();
    //    }
    //}

    public string GetNote(string? note, string? noteOderPunkte, string? punkte, string fach, string tendenz)
    {
        if (note == "84" || note == "84.00" || note == "A" || note == "Attest") return "AT";
        if (note == "99") return "NB";

        if (noteOderPunkte == "P") return note + tendenz;
        return noteOderPunkte == "N" ? note : "";
    }

    internal string GetNote(string jahrgang, List<dynamic>? marksPerLesson, string fach)
    {
        var linkeSeite = Nachname + ", " + Vorname + " (" + Klasse + "), " + fach;

        List<string> notenWebuntis = new List<string>();

        foreach (var x in marksPerLesson)
        {
            var dict = (IDictionary<string, object>)x;
            if (dict["Name"] != null && dict["Name"].ToString().Contains(Vorname))
            {
                if (dict["Name"].ToString().Contains(Nachname))
                {
                    if (dict["Klasse"].ToString() != null && dict["Klasse"].ToString().Contains(Klasse))
                    {
                        if (dict["Fach"] != null && dict["Fach"].ToString() == fach)
                        {
                            if (dict["Gesamtnote"] != null && dict["Gesamtnote"].ToString() != "")
                            {
                                notenWebuntis.Add(dict["Gesamtnote"].ToString());
                            }
                        }
                    }
                }
            }
        }

        List<string> notenSchild = new List<string>();

        foreach (var noteWebuntis in notenWebuntis.Distinct())
        {
            var result = "";

            switch (noteWebuntis)
            {
                case "15.0":
                    result = "1+";
                    break;
                case "14.0":
                    result = "1";
                    break;
                case "13.0":
                    result = "1-";
                    break;
                case "12.0":
                    result = "2+";
                    break;
                case "11.0":
                    result = "2";
                    break;
                case "10.0":
                    result = "2-";
                    break;
                case "9.0":
                    result = "3+";
                    break;
                case "8.0":
                    result = "3";
                    break;
                case "7.0":
                    result = "3-";
                    break;
                case "6.0":
                    result = "4+";
                    break;
                case "5.0":
                    result = "4";
                    break;
                case "4.0":
                    result = "4-";
                    break;
                case "3.0":
                    result = "5+";
                    break;
                case "2.0":
                    result = "5";
                    break;
                case "1.0":
                    result = "5-";
                    break;
                case "0.0":
                    result = "6";
                    break;
                case "84.0":
                case "A":
                case "Attest":
                    result = "AT";
                    break;
                case "99.0":
                    result = "NB";
                    break;
                case "97.0":
                    result = "AM";
                    break;
                case "77.0":
                case "T":
                    result = "E2";
                    break;
                case "78.0":
                    result = "E3";
                    break;
                default:
                    result = "";
                    break;
            }

            if (jahrgang.StartsWith("G") && !jahrgang.EndsWith("1"))
            {
                notenSchild.Add(result);
            }
            else
            {
                notenSchild.Add(result.Replace("+", "").Replace("-", ""));
            }
        }

        if (notenSchild.Distinct().Count() > 1)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            linkeSeite = (linkeSeite.PadRight(30, ' ')).Substring(0, 30) + ": Widersprechende Gesamtnoten: ";
            Global.ZeileSchreiben(linkeSeite + " Die zuletzt eingetragene Note gewinnt: " + notenSchild.Last(),
                string.Join(", ", notenSchild.Distinct()), ConsoleColor.Yellow,ConsoleColor.Gray);
            Console.ResetColor();
        }

        if (notenSchild.Count() == 0)
        {
            return "";
        }

        // Die zuletzt eingetragene Note gewinnt, falls .
        return notenSchild.Last();
    }





    public string GenerateMail()
    {
        if (string.IsNullOrEmpty(Nachname) || string.IsNullOrEmpty(Vorname) || string.IsNullOrEmpty(IdSchild)) return "";
        
        var id = IdSchild;

        if (!string.IsNullOrEmpty(ExterneIdNummer))
        {
            id = ExterneIdNummer;
        }

        return Bereinigen(Nachname.ToLower()).Substring(0, 1) + Bereinigen(Vorname.ToLower()).Substring(0, 1) +
               id.ToString().PadLeft(6, '0') + Global.MailDomain;
    }

    public string UmlauteBehandeln(string textinput)
    {
        string text = textinput;

        // deutsche Sonderzeichen
        text = Regex.Replace(text, "[æ|ä]", "ae");
        text = Regex.Replace(text, "[Æ|Ä]", "Ae");
        text = Regex.Replace(text, "[œ|ö]", "oe");
        text = Regex.Replace(text, "[Œ|Ö]", "Oe");
        text = Regex.Replace(text, "[ü]", "ue");
        text = Regex.Replace(text, "[Ü]", "Ue");
        text = Regex.Replace(text, "ß", "ss");

        // Sonderzeichen aus anderen Sprachen
        text = Regex.Replace(text, "[ã|à|â|á|å]", "a");
        text = Regex.Replace(text, "[Ã|À|Â|Á|Å]", "A");
        text = Regex.Replace(text, "[é|è|ê|ë]", "e");
        text = Regex.Replace(text, "[É|È|Ê|Ë]", "E");
        text = Regex.Replace(text, "[í|ì|î|ï]", "i");
        text = Regex.Replace(text, "[Í|Ì|Î|Ï]", "I");
        text = Regex.Replace(text, "[õ|ò|ó|ô]", "o");
        text = Regex.Replace(text, "[Õ|Ó|Ò|Ô]", "O");
        text = Regex.Replace(text, "[ù|ú|û|µ]", "u");
        text = Regex.Replace(text, "[Ú|Ù|Û]", "U");
        text = Regex.Replace(text, "[ý|ÿ]", "y");
        text = Regex.Replace(text, "[Ý]", "Y");
        text = Regex.Replace(text, "[ç|č]", "c");
        text = Regex.Replace(text, "[Ç|Č]", "C");
        text = Regex.Replace(text, "[Ð]", "D");
        text = Regex.Replace(text, "[ñ]", "n");
        text = Regex.Replace(text, "[Ñ]", "N");
        text = Regex.Replace(text, "[š]", "s");
        text = Regex.Replace(text, "[Š]", "S");

        return text;
    }

    public string Bereinigen(string textinput)
    {
        string text = textinput;

        text = text.ToLower(); // Nur Kleinbuchstaben
        text = UmlauteBehandeln(text); // Umlaute ersetzen


        text = Regex.Replace(text, "-", "_"); //  kein Minus-Zeichen
        text = Regex.Replace(text, ",", "_"); //  kein Komma            
        text = Regex.Replace(text, " ", "_"); //  kein Leerzeichen
        // Text = Regex.Replace(Text, @"[^\w]", string.Empty);   // nur Buchstaben

        text = Regex.Replace(text, "[^a-z]", string.Empty); // nur Buchstaben

        text = text.Substring(0, Math.Min(6, text.Length)); // Auf maximal 6 Zeichen begrenzen
        return text;
    }

    public string GetId(List<dynamic> schuelerZusatzdaten)
    {
        // Wenn der Schüler noch eine AtlantisID hat, dann wird die AtlantisID auf Webuntis gematcht.
        // Ansonsten wird die SchildID auf Webuntis gematcht.

        var externeIdNr = schuelerZusatzdaten.Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return Nachname == dict["Nachname"].ToString() &&
                   Vorname == dict["Vorname"].ToString() &&
                   Geburtsdatum == dict["Geburtsdatum"].ToString();
        }).Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Externe ID-Nr"].ToString();
        }).FirstOrDefault();

        return string.IsNullOrEmpty(externeIdNr) ? IdSchild : externeIdNr;
    }

    public void GetMaßnahmen(List<Datei?> dateien, List<string> maßnahmenString)
    {
        Massnahmen = new List<dynamic>();

        foreach (var datei in dateien)
        {
            if (!string.IsNullOrEmpty(datei.UnterordnerUndDateiname))
            {
                if (datei.UnterordnerUndDateiname.Contains("SchuelerVermerke"))
                {
                    foreach (var dateiZeile in datei)
                    {
                        foreach (var m in maßnahmenString)
                        {
                            var dict = (IDictionary<string, object>)dateiZeile;

                            if (dict["Vermerkart"] != null && dict["Vermerkart"].ToString().Contains(m))
                            {
                                if (dict["Geburtsdatum"].ToString() == Geburtsdatum &&
                                    dict["Nachname"].ToString() == Nachname &&
                                    dict["Vorname"].ToString() == Vorname)
                                {
                                    Massnahmen.Add(dateiZeile);
                                    AlleMaßnahmenUndVorgänge += dict["Datum"] + ":" + dict["Vermerkart"] + ", ";
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public string GetUrl(string v)
    {
        return "https://bkb.wiki/antraege_formulare:" + v + "?@Schüler*in@=" + Vorname + "_" + Nachname + "&@Klasse@=" +
               Klasse;
    }

    internal string GetWikiLink(string v, int stunden)
    {
        return @"\\ [[antraege_formulare:" + v + "?@Schüler*in@=" + Vorname + "_" + Nachname + "&@Klasse@=" + Klasse +
               "&@... unentschuldigte Fehlstunden:@=" + stunden + "]]";
    }

    internal string GetMaßnahmenAlsWikiLinkAufzählung()
    {
        var x = "";
        var bezeichnung = "";

        foreach (var item in Massnahmen.OrderBy(y => y.Datum))
        {
            var dict = (IDictionary<string, object>)item;

            if (dict["Vermerkart"].ToString().Contains("ahnung"))
            {
                bezeichnung = "Mahnung";
            }

            if (dict["Vermerkart"].ToString().Contains("ttestp"))
            {
                bezeichnung = "Attestpflicht";
            }

            if (dict["Vermerkart"].ToString().Contains("eilkonf") || dict["Vermerkart"].ToString().Contains("rdnungsm"))
            {
                bezeichnung = "Teilkonferenz";
            }

            if (bezeichnung == "")
            {
                bezeichnung = dict["Vermerkart"].ToString();
            }

            if (bezeichnung == "")
            {
                bezeichnung = dict["Vermerkart"].ToString();
            }

            //x += "[[" + bezeichnung + "]],";
            x += bezeichnung + " (" + dict["Datum"].ToString() + @")\\ ";
        }

        return x.TrimEnd(' ');
    }

    public void GetAbwesenheiten(List<Datei?> dateien, string id)
    {
        Abwesenheiten = new List<dynamic>();

        foreach (var datei in dateien)
        {
            if (!string.IsNullOrEmpty(datei.UnterordnerUndDateiname))
            {
                if (datei.UnterordnerUndDateiname.Contains("AbsencePerStudent"))
                {
                    foreach (var rec in datei)
                    {
                        var dict = (IDictionary<string, object>)rec;

                        if (dict["Externe Id"] != null && dict["Externe Id"].ToString() == id)
                        {
                            Abwesenheiten.Add(rec);
                        }
                    }
                }
            }
        }
    }

    public void GetF2(int verjährungUnbescholtene, int schonfrist)
    {
        // Für jede Abwesenheit ...
        foreach (var abwesenheit in Abwesenheiten)
        {
            // ... die nict entschuldigt ist ...
            if (abwesenheit.Status == "offen" || abwesenheit.Status == "nicht entsch.")
            {
                // ... sofern es noch keine Maßnahme gab in diesem Jahr ...
                if (JuengsteMassnahmeInDiesemSj == null)
                {
                    // ... und die Abwesenheit nicht bereits verjährt ist ...
                    if (DateTime.Parse(abwesenheit.Datum).Date.AddDays(verjährungUnbescholtene) > DateTime.Now.Date)
                    {
                        // ... und die Schonfrist abgelaufen ist ...
                        if (DateTime.Parse(abwesenheit.Datum).Date.AddDays(schonfrist) <= DateTime.Now.Date)
                        {
                            // ... und die Abwesenheit in dieses SJ fällt.
                            if (DateTime.Parse(abwesenheit.Datum) >
                                new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1))
                            {
                                var dict = (IDictionary<string, object>)abwesenheit;
                                F2 = F2 + Convert.ToInt32(dict["Fehlstd."]);
                            }
                        }
                    }
                }
            }
        }
    }

    public void GetJüngsteMaßnahmeInDiesemSj()
    {
        foreach (var maßnahme in Massnahmen)
        {
            var dict = (IDictionary<string, object>)maßnahme;

            if (DateTime.Parse(dict["Datum"].ToString()!) > new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1))
            {
                JuengsteMassnahmeInDiesemSj = maßnahme;
            }
        }
    }

    public void GetF3(int verjährungUnbescholtene, int schonfrist)
    {
        foreach (var abwesenheit in Abwesenheiten)
        {
            var dict = (IDictionary<string, object>)abwesenheit;

            if (dict["Status"] != null &&
                (dict["Status"].ToString() == "offen" || dict["Status"].ToString() == "nicht entsch."))
            {
                if (JuengsteMassnahmeInDiesemSj == null)
                {
                    if (DateTime.Parse(dict["Datum"].ToString()).Date.AddDays(verjährungUnbescholtene) >
                        DateTime.Now.Date)
                    {
                        if (DateTime.Parse(dict["Datum"].ToString()).Date.AddDays(schonfrist) > DateTime.Now.Date)
                        {
                            if (DateTime.Parse(dict["Datum"].ToString()) >
                                new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1))
                            {
                                F3 = F3 + Convert.ToInt32(dict["Fehlstd."]);
                            }
                        }
                    }
                }
            }
        }
    }

    public void GetF2PlusF3(int verjährungUnbescholtene, int schonfrist)
    {
        foreach (var a in Abwesenheiten)
        {
            var dict = (IDictionary<string, object>)a;
            if (dict["Status"] != null &&
                (dict["Status"].ToString() == "offen" || dict["Status"].ToString() == "nicht entsch."))
            {
                if (JuengsteMassnahmeInDiesemSj == null)
                {
                    if (DateTime.Parse(dict["Datum"].ToString()).Date.AddDays(verjährungUnbescholtene) >=
                        DateTime.Now.Date)
                    {
                        if (DateTime.Parse(dict["Datum"].ToString()).Date >
                            new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1))
                        {
                            F2PlusF3 = F2PlusF3 + Convert.ToInt32(dict["Fehlstd."].ToString());
                        }
                    }
                }
            }
        }
    }

    public void GetF2M(int verjährungUnbescholtene, int schonfrist)
    {
        foreach (var a in Abwesenheiten)
        {
            var dict = (IDictionary<string, object>)a;

            if (dict["Status"] != null &&
                (dict["Status"].ToString() == "offen" || dict["Status"].ToString() == "nicht entsch."))
            {
                // Falls es schon eine Maßnahme gab ...
                if (JuengsteMassnahmeInDiesemSj != null)
                {
                    var dictJ = (IDictionary<string, object>)JuengsteMassnahmeInDiesemSj;

                    // zählen alle n.e. Fehlzeiten seit Maßnahme    
                    if (DateTime.Parse(dict["Datum"].ToString()).Date > DateTime.Parse(dictJ["Datum"].ToString()))
                    {
                        if (DateTime.Parse(dict["Datum"].ToString()).Date.AddDays(schonfrist) <= DateTime.Now.Date)
                        {
                            if (DateTime.Parse(dict["Datum"].ToString()) >
                                new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1))
                            {
                                F2M = F2M + Convert.ToInt32(dict["Fehlstd."].ToString());
                            }
                        }
                    }
                }
            }
        }
    }

    public void GetF2MplusF3()
    {
        foreach (var a in Abwesenheiten)
        {
            var dict = (IDictionary<string, object>)a;

            if (dict["Status"] != null &&
                (dict["Status"].ToString() == "offen" || dict["Status"].ToString() == "nicht entsch."))
            {
                // Falls es schon eine Maßnahme gab ...
                if (JuengsteMassnahmeInDiesemSj != null)
                {
                    var dictJ = (IDictionary<string, object>)JuengsteMassnahmeInDiesemSj;
                    // zählen alle n.e. Fehlzeiten seit Maßnahme
                    if (DateTime.Parse(dict["Datum"].ToString()).Date > DateTime.Parse(dictJ["Datum"].ToString()))
                    {
                        if (DateTime.Parse(dict["Datum"].ToString()).Date >
                            new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1))
                        {
                            F2MplusF3 = F2MplusF3 + Convert.ToInt32(dict["Fehlstd."].ToString());
                        }
                    }
                }
            }
        }
    }

    public void ZieldateiSpeichern(string art, string datum, string dateiName)
    {
        PdfDocument quelldatei = PdfReader.Open(dateiName, PdfDocumentOpenMode.Import);
        PdfDocument zieldatei = new PdfDocument();

        foreach (var pdfSeite in PdfSeiten)
        {
            zieldatei.AddPage(quelldatei.Pages[pdfSeite.Seite - 1]);
        }

        zieldatei.Save(Zielordner + "/" + $"{Nachname}_{Vorname}_{Klasse}_{art}_{datum}.pdf");
        Global.ZeileSchreiben(" -> " + $"{Nachname}_{Vorname}_{Klasse}_{art}_{datum}.pdf", "erstellt", ConsoleColor.Yellow,ConsoleColor.Gray);
    }

    public string GetJahrgang(List<dynamic>? schuelerBasisdaten)
    {
        var jahrgang = schuelerBasisdaten.Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Vorname"].ToString() == Vorname && dict["Nachname"].ToString() == Nachname &&
                   dict["Geburtsdatum"].ToString() == Geburtsdatum;
        }).Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Jahrgang"].ToString();
        }).FirstOrDefault();

        return string.IsNullOrEmpty(jahrgang) ? "" : jahrgang;
    }

    public string GetJahrgang(List<dynamic>? klassen, string jahr, DateTime zeugnisdatum, IDictionary<string, object> dictBasis)
    {
        var jahrgang = klassen
            .Where(record =>
            {
                var dict = (IDictionary<string, object>)record;
                return dict["InternBez"].ToString() == this.Klasse;
            })
            .Select(record =>
            {
                var dict = (IDictionary<string, object>)record;
                return dict["Jahrgang"].ToString();
            })
            .FirstOrDefault();

        if (jahrgang == "")
        {
            Console.WriteLine("Die Klasse scheint in Schild keinen Jahrgang zu haben.");
        }

        try
        {
            int jahrHeute = Convert.ToInt32(Global.AktSj[0]);
            int differenz = jahrHeute - Convert.ToInt32(jahr);
            var jahrgangZahlHeute = Convert.ToInt32(jahrgang.Split('-').Last());
            var jahrgangZahlDamals = jahrgangZahlHeute - differenz;
            return jahrgang.Replace(jahrgang.Split('-').Last(), jahrgangZahlDamals.ToString().PadLeft(2, '0'));
        }
        catch (Exception e)
        {
            return dictBasis["Jahrgang"].ToString();
        }
    }

    public IEnumerable<DateTime> GetZeugnisDatums(List<dynamic>? atlantiszeugnisse)
    {
        var dates = new List<DateTime>();

        foreach (var atlantiszeugnis in atlantiszeugnisse)
        {
            var dict = (IDictionary<string, object>)atlantiszeugnis;

            if (dict["Field1"].ToString().Replace("'","") == Nachname)
            {
                //var dat = Global.DMMyyyyToDatum(dict["Field3"].ToString());
                var dat = dict["Field3"].ToString().Replace("'","");

                if (dat == Geburtsdatum)
                {
                    if (!string.IsNullOrEmpty(dict["Field4"].ToString()))
                    {
                        var x = DateTime.ParseExact(dict["Field4"].ToString().Replace("'",""), "d.M.yyyy", CultureInfo.InvariantCulture);
                        if (!dates.Contains(x))
                        {
                            dates.Add(x);
                        }
                    }
                }
            }
        }

        return dates;
    }

    public dynamic GetErzieherart(string? art, string sorgeberechtigt, string geschlecht)
    {
        if (string.IsNullOrEmpty(art) || art == "0")
        {
            if (this.IstVolljaehrig())
            {
                if (geschlecht == "W")
                {
                    return "Schülerin ist volljährig";
                }
                else
                {
                    return "Schüler ist volljährig";
                }
            }

            return "";
        }

        if (art == "V")
        {
            return "Vater";
        }

        if (art == "M")
        {
            return "Mutter";
        }

        if (art == "E")
        {
            return "Eltern";
        }

        if (sorgeberechtigt == "J")
        {
            return "(sonst.) gesetzl. Vertreter";
        }

        return "Sonstige";
    }

    public bool IstVolljaehrig()
    {
        // Versuchen, das Datum zu parsen
        if (DateTime.TryParseExact(
                this.Geburtsdatum,
                "dd.MM.yyyy",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime geburtsdatumParsed))
        {
            // Berechnung des Alters
            DateTime heute = DateTime.Now;
            int alter = heute.Year - geburtsdatumParsed.Year;

            // Wenn der Geburtstag in diesem Jahr noch nicht war, Alter um 1 reduzieren
            if (heute < geburtsdatumParsed.AddYears(alter))
            {
                alter--;
            }

            // Volljährigkeit prüfen (ab 18 Jahren)
            return alter >= 18;
        }
        else
        {
            // Ungültiges Datum
            throw new ArgumentException("Das Geburtsdatum ist ungültig.");
        }
    }

    /// <summary>
    /// Wenn ein Unterricht in einem Halbjahr nicht stattfindet, wird er nicht in die Leistungsdaten übernmommen.
    /// Wenn ein Unterricht nach den Sommerferien nur ganz kurz unterrichtet wurde, 
    /// </summary>
    /// <param name="dict"></param>
    /// <returns></returns>
    public bool UnterrichtIstRelevantFürZeugnisInDiesemAbschnitt(IDictionary<string, object> dict)
    {
        var von = dict["startDate"].ToString() == ""
            ? new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1)
            : DateTime.Parse(dict["startDate"].ToString());
        var bis = dict["endDate"].ToString() == ""
            ? new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31)
            : DateTime.Parse(dict["endDate"].ToString());

        // Der Unterricht mit muss im aktuellen Abschnitt liegen, um berücksichtigt zu werden

        var ersterSchultag = new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1);
        var halbjahreswechsel = Global.Halbjahreszeugnisdatum;
        var letzterSchultag = new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31);

        // Der Unterricht muss mindestens teilweise in den interessierenden Abschnitt fallen
        if (Global.Abschnitt == "1")
        {
            // Wenn der Unterricht irgendwann im ersten Hj beginnt, ist er relevant
            if (ersterSchultag <= von && von <= halbjahreswechsel)
            {
                var dauer = Math.Min((halbjahreswechsel - von).TotalDays, (bis - von).TotalDays);

                // Ein Unterricht muss länger als 30 Tage unterricht worden sein, um im Zeugnis aufzutauchen.
                // So wird Verwirrung zwischen den Kursen vermieden 
                if (dauer > 30)
                {
                    return true;
                }
                else
                {
                    // Wenn der Unterricht zu Beginn des Schuljahres weniger als 30 Tage stattgefunden hat, wird er ignorert,
                    // weil dann die Annahme gilt, dass der Schüler zu Beginn keinem / dem faschen Kurs zugeordnet war.

                    if (ersterSchultag.AddDays(30) < bis)
                    {
                        Console.WriteLine("   " + Nachname + ", " + Vorname + ", " + dict["subject"].ToString() + ", " +
                                          dict["startDate"].ToString() + "-" + dict["endDate"].ToString() +
                                          ": zu wenige Tag für eine Berücksichtigung im Zeugnis");
                        return true;
                    }
                    else
                    {
                        // Wenn der Unterricht später im Jahr nur sehr kurz war, wird gefragt
                        Console.WriteLine("   " + Nachname + ", " + Vorname + ", " + dict["subject"].ToString() + ", " +
                                          dict["startDate"].ToString() + "-" + dict["endDate"].ToString() +
                                          ": Was ist zu tun?");
                    }
                }
            }
        }

        if (Global.Abschnitt == "2")
        {
            // Wenn der Unterricht irgendwann im weiten Hj endet, ist er relevant
            if (halbjahreswechsel <= bis && bis <= letzterSchultag)
            {
                var dauer = Math.Min((letzterSchultag - von).Days, (bis - von).Days);

                // Ein Unterricht muss länger als 30 Tage unterricht worden sein, um im Zeugnis aufzutauchen.
                // So wird Verwirrung zwischen den Kursen vermieden 
                if (dauer > 30)
                {
                    return true;
                }
                else
                {
                    Console.WriteLine("   " + Nachname + ", " + Vorname + ", " + dict["subject"].ToString() + ", " +
                                      dict["startDate"].ToString() + "-" + dict["endDate"].ToString() +
                                      ": zu wenig Tag für eine Berücksichtigung im Zeugnis");
                    return false;
                }
            }
        }

        return false;
    }

    public string GetJahr(DateTime zeugnisdatum)
    {
        return zeugnisdatum.Month <= 8 ? zeugnisdatum.AddYears(-1).Year.ToString() : zeugnisdatum.Year.ToString();
    }

    public string GetAbschnitt(DateTime zeugnisdatum)
    {
        return zeugnisdatum.Month is <= 8 and > 2 ? "2" : "1";
    }

    public string GetIdSchildFromMail()
    {
        var match = MyRegex().Match(MailSchulisch);
        return match.Success ? match.Value.TrimStart('0') : "";
    }

    [GeneratedRegex(@"\d{1,6}")]
    private static partial Regex MyRegex();
}