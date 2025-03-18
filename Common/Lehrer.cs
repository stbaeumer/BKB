public class Lehrer
{
    public Lehrer()
    {
    }

    public int IdUntis { get; internal set; }
    public string Kürzel { get; internal set; }
    public string? Mail { get; internal set; }
    public string Geschlecht { get; internal set; }
    public double Deputat { get; internal set; }
    public string? Nachname { get; internal set; }
    public string? Vorname { get; internal set; }
    public string? Titel { get; internal set; }
    public string? Raum { get; internal set; }
    public DateTime Geburtsdatum { get; internal set; }
    public double AusgeschütteteAltersermäßigung { get; internal set; }
    public int ProzentStelle { get; internal set; }
    public int AlterAmErstenSchultagDiesesJahres { get; internal set; }
    public string? Flags { get; internal set; }
    public string Beschreibung { get; internal set; }
    public string? Text2 { get; internal set; }

    internal int GetAlterAmErstenSchultagDiesesJahres()
    {
        int years = Convert.ToInt32(Global.AktSj[0]) - Geburtsdatum.Year;
        DateTime birthday = Geburtsdatum.AddYears(years);
        if (new DateTime(2000 + Convert.ToInt32(Global.AktSj[0]), 7, 31).CompareTo(birthday) < 0)
        {
            years--;
        }

        return years;
    }

    internal int GetProzentStelle()
    {
        return Convert.ToInt32(Math.Floor(100 / 25.5 * Deputat));
    }

    internal void CheckAltersermäßigung(bool diesesJahr)
    {
        if (diesesJahr)
        {
            switch (AlterAmErstenSchultagDiesesJahres)
            {
                case >= 60:
                {
                    switch (ProzentStelle)
                    {
                        case >= 96 when AusgeschütteteAltersermäßigung != 3:
                        {
                            var x = " " + Kürzel + " war zu Beginn des SJ " + AlterAmErstenSchultagDiesesJahres + " alt und müsste mit einer " + ProzentStelle +
                                    "%-Stelle 3 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                            Console.WriteLine(Global.InsertLineBreaks(x, 77));
                            break;
                        }
                        case >= 75 and < 96 when AusgeschütteteAltersermäßigung != 2:
                        {
                            var y = " " + Kürzel + " war zu Beginn des SJ " + AlterAmErstenSchultagDiesesJahres + " alt und müsste mit einer " + ProzentStelle +
                                    "%-Stelle 2 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                            Console.WriteLine(Global.InsertLineBreaks(y, 77));
                            break;
                        }
                        case >= 50 and < 75 when AusgeschütteteAltersermäßigung != 1.5:
                        {
                            var z = " " + Kürzel + " war zu Beginn des SJ " + AlterAmErstenSchultagDiesesJahres + " alt und müsste mit einer " + ProzentStelle +
                                    "%-Stelle 1.5 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                            Console.WriteLine(Global.InsertLineBreaks(z, 77));
                            break;
                        }
                    }

                    break;
                }
                case >= 55 and < 60:
                {
                    switch (ProzentStelle)
                    {
                        case 100 when AusgeschütteteAltersermäßigung != 1:
                        {
                            var x = " " + Kürzel + " war zu Beginn des SJ " + AlterAmErstenSchultagDiesesJahres + " alt und müsste mit einer " + ProzentStelle +
                                    "%-Stelle 1 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                            Console.WriteLine(Global.InsertLineBreaks(x, 77));
                            break;
                        }
                        case >= 50 and < 100 when AusgeschütteteAltersermäßigung != 0.5:
                        {
                            var y = Kürzel + " war zu Beginn des SJ " + AlterAmErstenSchultagDiesesJahres + " alt und müsste mit einer " + ProzentStelle +
                                    "%-Stelle 0,5 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                            Console.WriteLine(Global.InsertLineBreaks(y, 77));
                            break;
                        }
                    }
                    break;
                }
            }
        }
        else
        {
            if (AlterAmErstenSchultagDiesesJahres + 1 >= 60)
            {
                switch (ProzentStelle)
                {
                    case >= 96 when AusgeschütteteAltersermäßigung != 3:
                    {
                        var x = " " + Kürzel + " wird vor Beginn des kommenden SJ " + (AlterAmErstenSchultagDiesesJahres + 1) + " geworden sein und muss mit einer " + ProzentStelle +
                                "%-Stelle 3 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                        Console.WriteLine(Global.InsertLineBreaks(x, 77));
                        break;
                    }
                    case >= 75 and < 96 when AusgeschütteteAltersermäßigung != 2:
                    {
                        var x = " " + Kürzel + " wird vor Beginn des kommenden SJ " + (AlterAmErstenSchultagDiesesJahres + 1) + " geworden sein und muss mit einer " + ProzentStelle +
                                "%-Stelle 2 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                        Console.WriteLine(Global.InsertLineBreaks(x, 77));
                        break;
                    }
                    case >= 50 and < 75 when AusgeschütteteAltersermäßigung != 1.5:
                    {
                        var x = " " + Kürzel + " wird vor Beginn des kommenden SJ " + (AlterAmErstenSchultagDiesesJahres + 1) + " geworden sein und muss mit einer " + ProzentStelle +
                                "%-Stelle 1.5 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                        Console.WriteLine(Global.InsertLineBreaks(x, 77));
                        break;
                    }
                }
            }

            switch (AlterAmErstenSchultagDiesesJahres + 1)
            {
                case >= 55 when AlterAmErstenSchultagDiesesJahres < 60:
                {
                    if (ProzentStelle == 100 && AusgeschütteteAltersermäßigung != 1)
                    {
                        var x = " " + Kürzel + " wird vor Beginn des kommenden SJ " + (AlterAmErstenSchultagDiesesJahres + 1) + " geworden sein und muss mit einer " + ProzentStelle +
                                "%-Stelle 1 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                        Console.WriteLine(Global.InsertLineBreaks(x, 77));
                    }

                    if (ProzentStelle >= 50 && ProzentStelle < 100 && AusgeschütteteAltersermäßigung != 0.5)
                    {
                        var x = " " + Kürzel + " wird vor Beginn des kommenden SJ " + (AlterAmErstenSchultagDiesesJahres + 1) + " geworden sein und muss mit einer " + ProzentStelle +
                                "%-Stelle 0,5 Std Ermäßigung bekommen statt wie bisher : " + AusgeschütteteAltersermäßigung + " Std.";
                        Console.WriteLine(Global.InsertLineBreaks(x, 77));
                    }

                    break;
                }
            }
        }
    }
}