package gcal2excel;

public class Gcal2Excel
{

    public Gcal2Excel()
    {
        //
    }

    public static void main(String args[])
    {
        if (args.length != 6)
        {
            System.err.println("Usage: Gcal2Excel email password calendar_id"
                    + " start_date(yyyy-mm-dd) end_date(yyyy-mm-dd)"
                    + " output_filename.xls");
            System.exit(1);
        }

        Converter converter = new Converter(args[0], args[1], args[2]);

        boolean ok = converter.convert(args[3], args[4], args[5]);
        
        if(ok)
        {
            System.out.println("OK!");
        }
        else
        {
            System.out.println("ERROR!");
        }

        System.err.println(converter.getErrorMessage());
        //System.out.println(converter.getDebugMessage());
    }
}