using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;

public class TrimAndReplaceUnderscoreConverter : DefaultTypeConverter
{
    public override object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
    {
        if (text == null)
        {
            return base.ConvertFromString(text, row, memberMapData);
        }

        // Entferne führende und nachfolgende Leerzeichen und ersetze Unterstriche
        return text.Trim().Replace("_", "");
    }
}