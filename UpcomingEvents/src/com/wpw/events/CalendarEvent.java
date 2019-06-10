package com.wpw.events;

import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;

/*******************************************************************************
 * Represents a single calendar event (which can span multiple days)
 * 
 * @author  Wayne Wooten
 * @version 1.0
 * @since   2017-12-14
 *
 *******************************************************************************/
public class CalendarEvent {
	public static final DateTimeFormatter DATE_FORMATTER1 = DateTimeFormatter.ofPattern("M/d");
	public static final DateTimeFormatter DATE_FORMATTER2 = DateTimeFormatter.ofPattern("d");
	
	private static boolean wordFormat = true;
	
	private LocalDate startDate;
	private LocalDate endDate;
	private String description;
	private boolean multiDay;
	
	/**
	 * Class constructor for an all day event
	 * 
	 * @param startDate
	 *     the date of the event
	 * @param description
	 *     the event description
	 */
	public CalendarEvent(LocalDate startDate, String description) {
		this.startDate   = startDate;
		this.endDate     = startDate;
		this.description = description;
	}
	
	/**
	 * Class constructor for an event at a specific time
	 * 
	 * @param startDate
	 *     the date of the event
	 * @param startTime
	 *     the time of the event
	 * @param description
	 *     the event description
	 */
	public CalendarEvent(LocalDate startDate, LocalTime startTime, String description) {
		this.startDate   = startDate;
		this.endDate     = startDate;
		this.description = description;
	}
	
	/**
	 * Sets a new end date for the event - only works for events on consecutive dates
	 * 
	 * @param endDate
	 *     the end date of the event
	 */
	public void setEndDate(LocalDate endDate) {
		if (isNextDay(endDate)) {
			this.endDate = endDate;
			multiDay = true;
		}
	}
	
	/**
	 * Determines if the given date is the day after the end date of the event
	 * 
	 * @param date
	 *     the date in question
	 * @return
	 *     true if the given date is the day after the end date of the event
	 */
	public boolean isNextDay(LocalDate date) {
		return endDate.equals(date.minusDays(1));
	}
	
	/**
	 * @return
	 *     the current event in string format
	 */
	@Override
	public String toString() {
		return String.format("%s %s %s", getDateString(), getDash(), description);
	}
	
	/**
	 * Gets the event date in string format
	 * 
	 * @return
	 *     the event date in string format
	 */
	public String getDateString() {
		if (multiDay) {
			if (startDate.getMonth().equals(endDate.getMonth())) {
				return startDate.format(DATE_FORMATTER1) + "-" + endDate.format(DATE_FORMATTER2);
			} else {
				return startDate.format(DATE_FORMATTER1) + "-" + endDate.format(DATE_FORMATTER1);
			}
			
		} else {
			return startDate.format(DATE_FORMATTER1);
		}
	}
	
	/**
	 * Gets the event description
	 * 
	 * @return
	 *     the event description
	 */
	public String getDescription() {
		return description;
	}
	
	/**
	 * Determines if output characters should be for a MS Word document or standard out
	 * 
	 * @param wordFormat
	 *     the boolean parameter to determine the format of output characters
	 */
	public static void setWordFormat(boolean wordFormat) {
		CalendarEvent.wordFormat = wordFormat;
	}

	/**
	 * Gets the dash formated for MS Word or standard out
	 * 
	 * @return
	 *     the dash formated for MS Word or standard out
	 */
	private static String getDash() {
		return (wordFormat) ? "\u2013" : "-";
	}
}
