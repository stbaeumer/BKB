using Microsoft.Extensions.Configuration;
using PdfSharp.Pdf.Security;
using PdfSharp.Pdf;
using System.Text.RegularExpressions; // Add this directive for PdfDocument

public class PdfDateien : List<PdfDatei>
{
    public List<string> Muster { get; set; }
    public string? InputFolder { get; private set; }
    public string? OutputFolder { get; private set; }

    public PdfDateien()
    {
    }

    public PdfDateien(Students students, List<string> suchmuster)
    {
        Global.DisplayHeader(Global.Header.H3, "Dateien einlesen:", Global.Protokollieren.Nein);
        
        foreach (string dateiName in Directory.GetFiles(Global.InputFolder, "*.pdf"))
        {
            var pdfDatei = new PdfDatei(dateiName);
            pdfDatei.Seiten.Read(dateiName);
            pdfDatei.Students = pdfDatei.GetStudentsMitSeiten(students);
            
            foreach (var student in pdfDatei.Students)
            {
                string art = student.PdfSeiten.GetArt(suchmuster);
                string datum = student.PdfSeiten.GetDatum();
                student.CreateFolderPdfDateien();
                student.ZieldateiSpeichern(art, datum, dateiName);
            }
            
            if(pdfDatei.Students.Any())
                pdfDatei.SeitenAusQuelldateienLöschen();
        }
    }

    public PdfDateien(Lehrers lehrers, IConfiguration configuration)
    {
        Global.DisplayHeader(Global.Header.H3, "Dateien einlesen:", Global.Protokollieren.Nein);
                
        var dateiPfad = Directory.GetFiles(Global.PfadExportdateien, "*.pdf").OrderByDescending(File.GetLastWriteTime).FirstOrDefault();

        if (dateiPfad != null)
        {
            var pdfDatei = new PdfDatei(dateiPfad);
            
            pdfDatei.SeitenEinlesen(dateiPfad);

            foreach (var seite in pdfDatei.Seiten)
            {   
                if (seite.Inhalt == null)
                {
                    continue;
                }

                foreach (var leh in lehrers)               
                {
                    if(!string.IsNullOrEmpty(leh.Mail))
                    {
                        if(IstValideEmail(leh.Mail))
                        {
                            if(seite.Inhalt.ToLower().Contains(leh.Mail.ToLower()))
                            {
                                var mail = new Mail();
                                var seitePdf = new PdfDocument();

                                // Extrahiere die Seite mit der Seitenzahl `s`
                                var pdfPage = pdfDatei.Pages[seite.Seite - 1]; // Seitenzahl ist 1-basiert, daher -1
                                seitePdf.AddPage(pdfPage);

                                // Speichere die extrahierte Seite in einer neuen PDF-Datei
                                var seitePdfName = Path.Combine(Global.PfadExportdateien, Path.GetFileNameWithoutExtension(dateiPfad) + "-" + leh.Kürzel + ".pdf");
                                seitePdf.Save(seitePdfName);

                                // Sende die extrahierte Seite per E-Mail
                                mail.Senden("Betreff: PDF-Seite", Global.SmtpUser, "Hier ist die angeforderte Seite.", seitePdfName, leh.Mail);
                            }                            
                        }
                    }
                }
            }
        }
    }

    private bool IstValideEmail(string email)
    {
        var emailRegex = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
        return Regex.IsMatch(email, emailRegex);
    }

    private PdfSharp.Pdf.PdfPage ConvertToPdfPage(PdfSeite seite)
{
    var pdfPage = new PdfSharp.Pdf.PdfPage();
    // Kopieren Sie hier die relevanten Inhalte von `PdfSeite` nach `PdfPage`
    // Beispiel: pdfPage.Contents = seite.Contents;
    return pdfPage;
}

    public void KennwortSetzen(IConfiguration configuration)
    {
        try
        {
            var fileGroupPdf =
                (from f in Directory.GetFiles(Global.PfadExportdateien, "*.pdf")
                    where !f.Contains("-Kennwort")
                    select f).ToList();

            Global.ZeileSchreiben("Dateien bereit für die Verschlüsselung", fileGroupPdf.Count == 0 ? "keine gefunden" : fileGroupPdf.Count.ToString(), ConsoleColor.White, ConsoleColor.DarkBlue);
            
            foreach (var file in fileGroupPdf)
            {
                Global.ZeileSchreiben(file, "bereit zum Erstellen einer kennwortgeschützten Kopie", ConsoleColor.Blue, ConsoleColor.Black);
            }

            if (fileGroupPdf.Count == 0)
            {
                return;
            }

            Console.WriteLine("");
            Console.WriteLine("   1: Schips");
            Console.WriteLine("   2: Notenlisten");
            Console.WriteLine("   3: andere PDFs");

            Global.Konfig("SchipsOderZeugnisseOderAnderePdfs", configuration, "Was wollen Sie verschlüsseln?");

            var passwort = "";
            var url = "";
            List<string> regex = new List<string>(); 

            if (Global.SchipsOderZeugnisseOderAnderePdfs == "1")
            {   
                Global.Konfig("SchipsUrl", configuration, "Schips-Url angeben");
                Global.Konfig("SchipsPasswort", configuration, "Schips-Kennwort festlegen");    
                passwort = Global.Verschluesseln(Global.SchipsPasswort);
                url = Global.SchipsUrl;
                regex.Add("schips");                
            }
            else if (Global.SchipsOderZeugnisseOderAnderePdfs == "2")
            {
                Global.Konfig("ZeugnisUrl", configuration, "Zeugnis-Url angeben");
                Global.Konfig("ZeugnisPasswort", configuration, "Zeugnis-Kennwort festlegen");    
                passwort = Global.Verschluesseln(Global.ZeugnisPasswort);
                url = Global.ZeugnisUrl;

            }
            else if (Global.SchipsOderZeugnisseOderAnderePdfs == "3")
            {
                passwort = Global.Entschluesseln(configuration["PdfKennwort"]);
            }
                
            foreach (string fileName in fileGroupPdf)
            {
                // Nurt wenn der Dateiname den Regex groß oder klein enthält.
                // Falls kein Regex, dann alle Dateien
                if (regex.Count == 0 || regex.Any(r => fileName.IndexOf(r, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(fileName);
                    PdfSecuritySettings securitySettings = document.SecuritySettings;
                    securitySettings.UserPassword = Global.Entschluesseln(passwort);
                    securitySettings.OwnerPassword = Global.Entschluesseln(passwort);
                    //securitySettings.PermitAccessibilityExtractContent = false;
                    securitySettings.PermitAnnotations = false;
                    securitySettings.PermitAssembleDocument = false;
                    securitySettings.PermitExtractContent = false;
                    securitySettings.PermitFormsFill = true;
                    securitySettings.PermitFullQualityPrint = false;
                    securitySettings.PermitModifyDocument = true;
                    securitySettings.PermitPrint = false;

                    var neueDatei = fileName.Replace(Path.GetFileNameWithoutExtension(fileName),
                        Path.GetFileNameWithoutExtension(fileName) + "-Kennwort");

                    document.Save(neueDatei);

                    Global.ZeileSchreiben(neueDatei, "Kopie mit Kennwort erstellt", ConsoleColor.Blue,ConsoleColor.Black);
                }                
            }
            Global.OpenWebseite(url);

        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
    }
}