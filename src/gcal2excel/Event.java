package gcal2excel;


public class Event
{
  private String task;
  private double hours;
  private String starts;
  private String ends;
  private String notes;
  private String location;

  public Event(String task, double hours, String starts, String ends, String notes, String location)
  {
  	this.task   = task;
  	this.hours  = hours;
  	this.starts = starts;
  	this.ends   = ends;
  	this.notes  = notes;
    this.location = location;
  }

  public String getTask()
  {
  	return this.task;
  }
  public double getHours()
  {
  	return this.hours;
  }
  public String getStarts()
  {
  	return this.starts;
  }
  public String getEnds()
  {
  	return this.ends;
  }
  public String getNotes()
  {
	return this.notes;
  }
  public String getLocation()
  {
	return this.location;
  }
}