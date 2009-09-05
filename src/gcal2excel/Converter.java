package gcal2excel;

import com.google.gdata.client.Query;
import com.google.gdata.client.calendar.CalendarQuery;
import com.google.gdata.client.calendar.CalendarService;
import com.google.gdata.data.DateTime;
import com.google.gdata.data.calendar.CalendarEventEntry;
import com.google.gdata.data.calendar.CalendarEventFeed;
import com.google.gdata.data.extensions.When;
import com.google.gdata.util.ServiceException;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.data.*;
import com.google.gdata.data.extensions.Where;

import com.google.gdata.data.PlainTextConstruct;

import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Vector;
import java.io.File;
import java.util.Iterator;
import java.util.ListIterator;

import jxl.*;
import jxl.write.*;
import jxl.write.WriteException;

/**
 * Creates an excel file from google calendar
 * using Google Calendar API and JExcel API
 */
public class Converter
{
  // The base URL for a user's calendar metafeed (needs a username appended).
  private final String METAFEED_URL_BASE =
      "http://www.google.com/calendar/feeds/";

  private final String SINGLE_FEED_URL_SUFFIX = "/private/full";

  // The URL for the event feed of the specified user's primary calendar.
  // (e.g. http://www.googe.com/feeds/calendar/calendar-id/private/full)
  private static URL singleFeedUrl = null;


  private final long MILISECONDS_IN_HOUR = 60*60*1000;

  private String email;
  private String password;
  private String calendarId;

  private String errorMessage;
  private String debugMessage;

 public Converter(String email, String password, String calendarId)
 {
	 this.email         = email ;
	 this.password      = password;
	 this.calendarId    = calendarId;
     this.errorMessage  = "";
     this.debugMessage  = "";
  }

  public boolean convert(String starts, String ends, String fileName)
  {
    Vector<Event> events = null;

    events = this.getCalendarData(this.email, this.password, this.calendarId, starts, ends);

    if(events==null) { return false; }

    return writeExcel(events, fileName);
  }

  private boolean writeExcel(Vector<Event> events, String fileName)
  {
  	try
	{
		WritableWorkbook workbook = Workbook.createWorkbook(new File(fileName));
		WritableSheet sheet = workbook.createSheet("Time Sheet", 0);

		WritableFont labelFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat labelFormat = new WritableCellFormat (labelFont);

		Label labelTask = new Label(0, 0, "Task", labelFormat);
		sheet.addCell(labelTask);

		Label labelHours = new Label(2, 0, "Hours", labelFormat);
		sheet.addCell(labelHours);

		Label labelStart = new Label(4, 0, "Start", labelFormat);
		sheet.addCell(labelStart);

		Label labelEnds = new Label(6, 0, "Ends", labelFormat);
		sheet.addCell(labelEnds);

		Label labelNotes = new Label(8, 0, "Notes", labelFormat);
		sheet.addCell(labelNotes);

		Label labelLocation = new Label(10, 0, "Location", labelFormat);
		sheet.addCell(labelLocation);

		int row = 2;
		ListIterator iterator = events.listIterator();
		while (iterator.hasNext())
		{
		    Event event = (Event)iterator.next();

		    jxl.write.Label task = new jxl.write.Label(0, row, event.getTask());
            sheet.addCell(task);


		    jxl.write.Number hours = new jxl.write.Number(2, row, event.getHours());
            sheet.addCell(hours);

                    jxl.write.Label starts = new jxl.write.Label(4, row, event.getStarts());
                    sheet.addCell(starts);

                    jxl.write.Label ends = new jxl.write.Label(6, row, event.getEnds());
                    sheet.addCell(ends);

                    jxl.write.Label comments = new jxl.write.Label(8, row, event.getNotes());
                    sheet.addCell(comments);

                    jxl.write.Label location = new jxl.write.Label(10, row, event.getLocation());
                    sheet.addCell(location);

                    row++;
		}

		Label labelTotal = new Label(0, (row+2), "Total Hours Billed: ", labelFormat);
		sheet.addCell(labelTotal);

		jxl.write.Formula formula = new jxl.write.Formula( 2, (row+2), "SUM(C3:C"+row+")" );
		sheet.addCell(formula);


	    workbook.write();
		workbook.close();

		return true;
	}
	catch(IOException e)
	{
      this.logError(Errors.FILE_WRITE_ERROR);
      return false;
	}
	catch(jxl.write.WriteException e)
	{
	  this.logError(Errors.EXCEL_WRITE_ERROR);
      return false;
	}
  }

  private Vector<Event> getCalendarData(
  		String email, String password,
 		String calendarId,
  		String starts, String ends)
  {


  	CalendarService myService = new CalendarService("anupomsyam-Gcal2ExcelWeb-v2");

    Vector<Event> events = null;

  	// Create the necessary URL objects.
  	try
    {
      singleFeedUrl = new URL(METAFEED_URL_BASE + calendarId + SINGLE_FEED_URL_SUFFIX);
      this.logDebug(singleFeedUrl.toString());
    }
    catch (MalformedURLException e)
    {
      // Bad URL
      this.logError(Errors.INVALID_URL_ERROR);
      return null;
    }

    try
    {
      myService.setUserCredentials(email, password); 
    }
    catch(AuthenticationException e)
    {
        this.logError(Errors.WRONG_CREDENTIAL_ERROR);
        return null;
    }
    
    try {
        events = dateRangeQuery(myService,
		      		DateTime.parseDate(starts),
		      		DateTime.parseDate(ends));
    }
    catch (IOException e)
    {
      // Communications error
      this.logError(Errors.COMMUNICATION_ERROR);
      return null;
    }
    catch (ServiceException e)
    {
      // Server side error
      this.logError(Errors.WRONG_CID_ERROR);
      return null;
    }

    return events;
  }

  /**
   * Prints the titles and start and end times of all the events.
   *
   * @param service An authenticated CalendarService object.
   * @param startTime Start time (inclusive) of events to print.
   * @param endTime End time (exclusive) of events to print.
   * @throws ServiceException If the service is unable to handle the request.
   * @throws IOException Error communicating with the server.
   */
  private Vector<Event> dateRangeQuery(CalendarService service,
      DateTime startTime, DateTime endTime) throws ServiceException,
      IOException
  {

    CalendarQuery myQuery = new CalendarQuery(singleFeedUrl);
    myQuery.setMinimumStartTime(startTime);
    myQuery.setMaximumStartTime(endTime);
    myQuery.addCustomParameter(new Query.CustomParameter("orderby", "starttime"));
    myQuery.addCustomParameter(new Query.CustomParameter("sortorder", "ascending"));
    myQuery.addCustomParameter(new Query.CustomParameter("singleevents", "true"));
    myQuery.addCustomParameter(new Query.CustomParameter("max-results", "10000"));

    // Send the request and receive the response:
    CalendarEventFeed resultFeed = service.query(myQuery, CalendarEventFeed.class);

    this.logDebug("Events from " + startTime.toString() + " to "
        + endTime.toString() + ":");
    this.logDebug("");

    int count = 1;
    Vector<Event> events = new Vector<Event>();

    for (int i = 0; i < resultFeed.getEntries().size(); i++)
    {
      CalendarEventEntry entry = resultFeed.getEntries().get(i);

      String title = entry.getTitle().getPlainText();
      java.util.List<When> eventTimes = entry.getTimes();

      String starts = null;
      String ends = null;
      String notes = null;
      String location = null;

      // Added by Kevin Jordan www.idreamincode.com
      notes  = ((PlainTextConstruct)(((TextContent)entry.getContent()).getContent())).getText();
      this.logDebug("plainText Content:\t" + notes);

      location  = ((Where)entry.getLocations().get(0)).getValueString() ;

      this.logDebug("plainText Content:\t" + location);
      // end of added section ------------

      double duration = 0;

      Iterator<When> iterator = eventTimes.iterator();

      while(iterator.hasNext())
      {
        When when = iterator.next();
      	starts = when.getStartTime().toUiString();
      	ends = when.getEndTime().toUiString();
      	duration = getWhenDiff(when);
      }


      events.addElement( new Event(title, duration, starts, ends, notes, location));
      this.logDebug("#"+count);
      this.logDebug("----------------------");
      this.logDebug("\t" + starts + "\t" + title +"\t" + ends + "\t" + duration
              +"\tnotes:" + notes  + "\tlocation:" + location );
      this.logDebug("----------------------");
      count++;

   }
   this.logDebug("");

   return events;
  }

  private double getWhenDiff(When when)
  {
  	 double hours = 0.0;

  	 try
  	 {
  	 	long diff  = when.getEndTime().getValue()-when.getStartTime().getValue();
  	 	if(diff>0)
  	 	{
  	 		hours = (double) diff/(MILISECONDS_IN_HOUR);
  	 	}
  	 }
  	 catch (Exception e)
  	 {
  	 	this.logError(Errors.TIME_DIF_ERROR);
  	 }
  	 return hours;
  }

  private void logError(String msg) {
    this.errorMessage += msg + "\n";
  }

  private void logDebug(String msg) {
    this.debugMessage += msg + "\n";
  }
  
  public String getErrorMessage()
  {
      return this.errorMessage;
  }
  
  public String getDebugMessage()
  {
      return this.debugMessage;
  }
}
