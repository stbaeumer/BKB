using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Platform;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia;
using Avalonia.Layout;
using Avalonia.Media.Imaging;

namespace BKB;

public partial class MainWindow : Window
{
    private Grid _hinweisContainer;
    public List<Border> Hinweise = new List<Border>();

    private const int HinweisHoehe = 15; // Fixe Höhe für Konsistenz
    private const int Abstand = 5; // Einheitlicher Abstand zwischen Hinweisen

    public Klassen Klassen { get; set; }
    public Students Students { get; set; }
    public Dateien Dateien { get; set; }

    public MainWindow()
    {
        InitializeComponent();
        
        _hinweisContainer = this.FindControl<Grid>("HinweisContainer");
       
        var documentsFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        var configPath = Path.Combine(documentsFolderPath, "BKB.json");

        var configuration = Global.GetConfiguration();
        
        // Sidebar Schuljahr
        var schuljahr = configuration["AktSj"];
        cbxAktSj.ItemsSource = new List<string> { "2023", "2024", "2025", "2026", "2027", "2028" };
        if (string.IsNullOrEmpty(schuljahr))
        {
            cbxAktSj.SelectedItem = Global.AktSj[0];
        }
        else
        {
            cbxAktSj.SelectedItem = schuljahr;
        }

        // Sidebar Abschnitt
        var abschnitt = configuration["Abschnitt"];
        cbxAbschnitt.ItemsSource = new List<string> { "1", "2" };
        if (string.IsNullOrEmpty(abschnitt))
        {
            cbxAbschnitt.SelectedItem = DateTime.Now.Month > 1 && DateTime.Now.Month <= 8 ? "2" : "1";
        }
        else
        {
            cbxAbschnitt.SelectedItem = abschnitt;
        }

        // Sidebar PfadExportdateien
        var pfadExportdateien = configuration["PfadExportdateien"];
        if (string.IsNullOrEmpty(pfadExportdateien))
        {
            lblPfadExportdateien.Text =
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            Global.PfadExportdateien = lblPfadExportdateien.Text;
        }
        else
        {
            lblPfadExportdateien.Text = pfadExportdateien;
            Global.PfadExportdateien = pfadExportdateien;
        }

        // Sidebar Datenaustausch
        var pfadSchilddateien = configuration["PfadSchilddateien"];
        if (string.IsNullOrEmpty(pfadSchilddateien))
        {
            lblPfadSchilddateien.Text =
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            Global.PfadSchilddateien = lblPfadSchilddateien.Text;
        }
        else
        {
            lblPfadSchilddateien.Text = pfadSchilddateien;
            Global.PfadSchilddateien = pfadSchilddateien;
        }


        // Sidebar MaxAlter
        var maxAlter = configuration["MaxDateiAlter"];
        if (string.IsNullOrEmpty(maxAlter))
        {
            tbxMaxAlter.Text = "3";
            Global.MaxDateiAlter = 3;
        }
        else
        {
            string input = String.Empty;
            try
            {
                int result = Int32.Parse(maxAlter);
                Global.MaxDateiAlter = result;
                tbxMaxAlter.Text = maxAlter;
            }
            catch (Exception)
            {
                HinweisErstellen("Kein zulässiger Wert", Brushes.AliceBlue);
            }
        }

        Dateien = new Dateien();
        Dateien.GetInteressierendeDateienMitAllenEigenschaften();
        var klasse = configuration["Klasse"];
    }

    private async void ImgClip_PointerPressed(object? sender, PointerPressedEventArgs e)
    {
        Global.Zeilen.Clear();
        HinweiseEntfernen();
        // Get top level from the current control. Alternatively, you can use Window reference instead.
        var topLevel = TopLevel.GetTopLevel(this);

        // Start async operation to open the dialog.
        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Dateien auswählen",
            AllowMultiple = true,
            FileTypeFilter = new List<FilePickerFileType>
            {
                new FilePickerFileType("erlaubt sind txt, csv und dat")
                {
                    Patterns = new[] { "*.csv", "*.dat", "*.txt" } // Nur csv und dat Dateien
                }
            }
        });

        var hinweis = "";

        var dateienImPfad = new List<string>();

        foreach (var file in files)
        {
            var localPath = new Uri(file.Path.ToString()).LocalPath;
            dateienImPfad.Add(localPath);
        }
        
        dateienImPfad = dateienImPfad.Distinct().ToList();
        
        Dateien.GetZeilen(dateienImPfad);
        
        // Prüfe, ob mindestens eine Datei existiert und Count > 0 ist
        if (Dateien.Any(d => d.Count > 0))
        {
            UploadImage.IsVisible = false; // StackPanel ausblenden
        }
        else
        {
            UploadImage.IsVisible = true; // StackPanel sichtbar lassen
            HinweisErstellen("Fehler! Es konnten keine Zeilen ausgelesen werden.", Brushes.Aqua);
            return;
        }

        ErstelleBilderGrid(BilderGrid, dateienImPfad);

        var klassen = new List<string>();
        klassen = Dateien.GetKlassen();

        cbxKlassen.Items.Clear();

        foreach (var klasse in klassen)
        {
            cbxKlassen.Items.Add(klasse);
        }

        if (cbxKlassen.Items.Count > 0)
        {
            cbxKlassen.SelectedIndex = 0; // Ersten Eintrag auswählen
            cbxKlassen.Background = new SolidColorBrush(Color.Parse("#B6FF85"));
            cbxKlassen.IsEnabled = true;
            
            if (Dateien.FirstOrDefault(x=>x.Name.ToLower().Contains("schuelerbasisdaten")).Count() == 0)
            {
                HinweisErstellen("Es muss zwingend eine aktuelle SchildSchuelerExport.txt hinzugefügt werden.", Brushes.Yellow);
            }
            else
            {
                // Die Schuelerbasisdaten müssen zwingend hochgeladen werden. 
                Students = new Students(@"ExportAusSchild/SchildSchuelerExport", "*.txt");
                Students = Students.DoppelteFiltern();
                if (!Students.IsAllesOk())
                {
                    return;
                }
                
                Students = new Students(Dateien);
                btnLosGehts.IsEnabled = true;    
            }
        }

        foreach (var zeile in Global.Zeilen)
        {
            HinweisErstellen(zeile.Item1 , Brushes.Yellow);
        }
    }

    private void Button_Click(object? sender, RoutedEventArgs e)
    {
        Global.Zeilen.Clear();
        HinweiseEntfernen();
        var iStudents = Students.GetInteressierende(cbxKlassen.SelectedItem.ToString());
        var iKlassen = iStudents.GetKlassen(new Klassen());

        var menue = MenueHelper.Einlesen(Dateien, Klassen, [], [], []);

        foreach (var eintrag in menue.Where(eintrag => eintrag.NotwendigeDateienDieNullOderEmptySind.Count > 0))
        {
            eintrag.Funktion(eintrag);
        }

        if (!menue.Any(eintrag => eintrag.NotwendigeDateienDieNullOderEmptySind.Count> 0))
        {
            HinweisErstellen("Nichts zu tun", Brushes.Yellow);
        }

        foreach (var (meldung, farbe) in Global.Zeilen)
        {
            HinweisErstellen(meldung, Brushes.Yellow);
        }
    }
    
    public enum HinweisFarbe
    {
        Gruen,
        Orange,
        Rot
    }

    public void HinweisErstellen(string text, IBrush farbe)
    {
        var rowIndex = Hinweise.Count; // Jede neue Zeile bekommt einen eindeutigen Index

        var border = new Border
        {
            Background = farbe,
            BorderBrush = farbe.ToString() == "Transparent" ? null : Brushes.White,
            BorderThickness = new Thickness(1),
            Width = 700,
            Height = HinweisHoehe,
            CornerRadius = new CornerRadius(3),
            Margin = new Thickness(0, 0, 0, Abstand), // Einheitlicher unterer Abstand
            Padding = new Thickness(10, 0, 10, 0),    // Innenabstand, um den Text vom Rand zu lösen
            Effect = new DropShadowEffect
            {
                Color = Colors.White,      // Schattenfarbe
                BlurRadius = 10,           // Weichzeichnung des Schattens
                Opacity = 0.5   
            }
        };


        var textBlock = new TextBlock
        {
            Text = text.Replace("................",""),
            TextWrapping = TextWrapping.Wrap,
            TextAlignment = TextAlignment.Left,
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Brushes.White,
            FontSize = 11,
            FontFamily = new FontFamily("Courier") // Schmale Monospace-Schrift
        };



        var stackPanel = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Center, // Inhalte vertikal zentrieren
            HorizontalAlignment = HorizontalAlignment.Stretch // Breite anpassen
        };

        stackPanel.Children.Add(textBlock);
        border.Child = stackPanel;

        border.PointerPressed += (sender, e) => HinweisEntfernen(border);

        // Neue Zeile im Grid hinzufügen
        _hinweisContainer.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        Grid.SetRow(border, rowIndex);

        _hinweisContainer.Children.Add(border);
        Hinweise.Add(border);
    }

    public void HinweiseEntfernen()
    {
        for (int i = 0; i < Hinweise.Count; i++)
        {
            HinweisEntfernen(Hinweise[i]);
        }
    }


    private void HinweisEntfernen(Border border)
    {
        int index = Hinweise.IndexOf(border);
        if (index == -1) return;

        Hinweise.RemoveAt(index);
        _hinweisContainer.Children.Remove(border);
        _hinweisContainer.RowDefinitions.RemoveAt(index);

        // Zeilen neu nummerieren
        for (int i = 0; i < Hinweise.Count; i++)
        {
            Grid.SetRow(Hinweise[i], i);
        }
    }


    private void HamburgerButton_Click(object? sender, RoutedEventArgs e)
    {
        HinweiseEntfernen();
        MenuSplitView.IsPaneOpen = !MenuSplitView.IsPaneOpen;
    }

    private void cbxAktSj_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        Global.Speichern("AktSj", cbxAktSj.SelectedItem.ToString());
    }

    private void Abschnitt_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        Global.Speichern("Abschnitt", cbxAbschnitt.SelectedItem.ToString());
    }
    
    private void TbxMaxAlter_TextChanged(object? sender, TextChangedEventArgs e)
    {
        var textBox = sender as TextBox;
        if (textBox != null)
        {
            string text = textBox.Text;

            // Überprüfen, ob der Text nur Ziffern enthält
            if (!string.IsNullOrEmpty(text) && text.All(char.IsDigit))
            {
                // Falls eine numerische Eingabe erfolgt, Text zurücksetzen
                textBox.Text = string.Concat(text.Where(char.IsDigit)); // Nur Ziffern bleiben
                Global.Speichern("MaxDateiAlter", text);
            }
        }
    }

    public void ErstelleBilderGrid(Grid gridContainer, List<string> dateien)
    {
        gridContainer.Children.Clear();
        gridContainer.RowDefinitions.Clear();
        gridContainer.ColumnDefinitions.Clear();

        int anzahl = Dateien.Count(x => x.Count() > 0);
        if (anzahl == 0) return;

        // Berechne die Spalten- und Zeilenanzahl für eine möglichst quadratische Form
        int spalten = (int)Math.Ceiling(Math.Sqrt(anzahl));
        int zeilen = (int)Math.Ceiling((double)anzahl / spalten);

        for (int i = 0; i < zeilen; i++)
            gridContainer.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        for (int j = 0; j < spalten; j++)
            gridContainer.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        // Größe abhängig von der Anzahl der Zeilen
        int imageSize = zeilen == 2 ? 48 : zeilen >= 3 ? 32 : 64;
        int fontSize = zeilen >= 3 ? 10 : 12;

        for (int index = 0; index < anzahl; index++)
        {
            var datei = Dateien.Where(x => x.Count > 0).ToList()[index];
            var dateiname = datei.Dateiname.Substring(0, Math.Min(16, datei.Dateiname.Length));
            var anzahlElemente = datei.Count();

            int row = index / spalten;
            int col = index % spalten;

            // StackPanel für Inhalt
            StackPanel stackPanel = new StackPanel
            {
                Orientation = Orientation.Vertical,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center };

            // Grid für Bilder
            Grid imageGrid = new Grid { Width = imageSize, Height = imageSize };

            Image image = new Image
            {
                Source = new Bitmap(AssetLoader.Open(new Uri("avares://gui/Assets/clip.png"))),
                Width = imageSize,
                Height = imageSize
            };

            imageGrid.Children.Add(image);

            TextBlock text1 = new TextBlock
            {
                Text = dateiname,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                FontSize = fontSize,
                Foreground = Brushes.White
            };

            TextBlock text2 = new TextBlock
            {
                Text = "\u2211 " + anzahlElemente.ToString(),
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                FontSize = fontSize,
                Foreground = Brushes.White
            };

            stackPanel.Children.Add(imageGrid);
            stackPanel.Children.Add(text1);
            stackPanel.Children.Add(text2);

            // Rahmen um das Element
            Border border = new Border
            {
                BorderBrush = Brushes.White,
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(5),
                Margin = new Thickness(5), // Abstand zwischen den Elementen
                Child = stackPanel,
                Effect = new DropShadowEffect
                {
                Color = Colors.Black,      // Schattenfarbe
                BlurRadius = 10,           // Weichzeichnung des Schattens
                Opacity = 0.5   
            }
            };

            // Event hinzufügen
            border.PointerPressed += ImgClip_PointerPressed;

            Grid.SetRow(border, row);
            Grid.SetColumn(border, col);
            gridContainer.Children.Add(border);
        }
    }

    private async void lblPfadSchilddateien_Click(object? sender, Avalonia.Input.PointerPressedEventArgs e)
    {
        var mainWindow = (Window)this.VisualRoot!; // Referenz zum Hauptfenster
        var dialog = new OpenFolderDialog { Title = "Bitte einen Ordner auswählen" };

        string? result = await dialog.ShowAsync(mainWindow); // Ordner auswählen
        if (!string.IsNullOrWhiteSpace(result))
        {
            lblPfadSchilddateien.Text = result;
            Global.Speichern("SchildDatenaustausch", result);
        }
    }
    private async void lblPfadExportdateien_Click(object? sender, Avalonia.Input.PointerPressedEventArgs e)
    {
        var mainWindow = (Window)this.VisualRoot!; // Referenz zum Hauptfenster
        var dialog = new OpenFolderDialog { Title = "Bitte einen Ordner auswählen" };

        string? result = await dialog.ShowAsync(mainWindow); // Ordner auswählen
        if (!string.IsNullOrWhiteSpace(result))
        {
            lblPfadExportdateien.Text = result;
            Global.Speichern("PfadExportdateien", result);
        }
    }
}