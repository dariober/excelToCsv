package excelToCsv;

import org.supercsv.prefs.CsvPreference;
import org.supercsv.quote.QuoteMode;
import org.supercsv.util.CsvContext;

public class NoQuoteMode implements QuoteMode
{
    @Override
    public boolean quotesRequired(String csvColumn, CsvContext context, CsvPreference preference)
    {
        return false;
    }
}