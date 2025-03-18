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
    
    public void KennwortSetzen()
    {
        try
        {
            var fileGroupPdf =
                (from f in Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\",
                        "*.pdf")
                    where !f.Contains("-Kennwort")
                    select f).ToList();

            Console.WriteLine("");
            foreach (var file in fileGroupPdf)
            {
                Global.ZeileSchreiben(file, "bereit zum Erstellen einer kennwortgeschützten Kopie", ConsoleColor.White, ConsoleColor.Green);
            }

            var documentsFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            var configPath = Path.Combine(documentsFolderPath, "BKB.json");

            var configuration = new ConfigurationBuilder()
                .AddJsonFile(configPath, optional: false, reloadOnChange: true).Build();
            var kennwort = configuration["Kennwort"];

            Console.WriteLine("");
            Console.WriteLine("   Bitte ein Kennwort wählen");
            Console.WriteLine("");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("      Ihr Kennwort [" + kennwort + "] : ");
            Console.WriteLine("");
            Console.ResetColor();

            var x = Console.ReadLine();

            if (x == "ö")
            {
                Global.OpenCurrentFolder();
            }

            if (x == "x")
            {
                Global.OpenWebseite("https://wiki.svws.nrw.de/mediawiki/index.php?title=Schnittstellenbeschreibung");
            }

            if ((x == "" && kennwort != ""))
            {
                Global.Speichern("Kennwort", kennwort!);
            }

            if (x != "")
            {
                Global.Speichern("Kennwort", x);
            }

            string[] fileGroupJpg =
                Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\", "*.png");

/*
            foreach (string fileName in (from f in fileGroupJpg where !f.Contains("-Kennwort") select f).ToList())
            {
                Document document = new Document(new Rectangle(288f, 144f), 10, 10, 10, 10);
                document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());

                using (var stream =
                       new FileStream(fileName + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    PdfWriter.GetInstance(document, stream);
                    document.Open();
                    using (var imageStream =
                           new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        var image = Image.GetInstance(imageStream);
                        image.SetAbsolutePosition(0, 0); // set the position to bottom left corner of pdf
                        image.ScaleAbsolute(iTextSharp.text.PageSize.A4.Height,
                            iTextSharp.text.PageSize.A4.Width); // set the height and width of image to PDF page size
                        document.Add(image);
                    }

                    document.Close();
                }
            }*/

            foreach (string fileName in fileGroupPdf)
            {
                PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(fileName);
                PdfSecuritySettings securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = kennwort!;
                securitySettings.OwnerPassword = kennwort!;
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

                Global.ZeileSchreiben(neueDatei, "Kopie mit Kennwort erstellt", ConsoleColor.Yellow,ConsoleColor.Gray);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
        finally
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("");
            Console.WriteLine("      Weiter mit ENTER");
            Console.WriteLine("");
            Console.ResetColor();

            Console.ReadKey();
        }
    }
}