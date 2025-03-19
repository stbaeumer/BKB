using Microsoft.Extensions.Configuration;
using PdfSharp.Pdf.Security;

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
            pdfDatei.PdfSeiten.Read(dateiName);
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

            if(fileGroupPdf.Count == 0){
                return;
            }

            var documentsFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            var configPath = Path.Combine(documentsFolderPath, "BKB.json");

            Global.Konfig("PdfKennwort", configuration, "PDF-Kennwort festlegen");

            foreach (string fileName in fileGroupPdf)
            {
                PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(fileName);
                PdfSecuritySettings securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = Global.PdfKennwort!;
                securitySettings.OwnerPassword = Global.PdfKennwort!;
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
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
    }
}