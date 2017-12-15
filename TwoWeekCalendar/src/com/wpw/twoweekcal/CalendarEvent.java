package com.wpw.twoweekcal;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.Month;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAdjusters;

/*******************************************************************************
 * Represents a single calendar event (which can span multiple days)
 * 
 * @author  Wayne Wooten
 * @version 1.0
 * @since   2017-12-14
 *
 *******************************************************************************/
public class CalendarEvent {
	public static final DateTimeFormatter DATE_FORMATTER1 = DateTimeFormatter.ofPattern("EE, MMM d");
	public static final DateTimeFormatter DATE_FORMATTER2 = DateTimeFormatter.ofPattern("MMM d");
	public static final DateTimeFormatter DATE_FORMATTER3 = DateTimeFormatter.ofPattern("d");
	public static final DateTimeFormatter TIME_FORMATTER1 = DateTimeFormatter.ofPattern("h:mm a");
	
	public static final DateTimeFormatter HEADER_DATE_FORMATTER = DateTimeFormatter.ofPattern("MMMM d");
	
	private static boolean includeAllDates = false;
	private static boolean wordFormat = true;
	
	private static LocalDate lastDatePrinted   = LocalDate.of(1970, Month.JANUARY, 1);
	private static LocalDate calendarStartDate = LocalDate.now().minusDays(1).with(TemporalAdjusters.next(DayOfWeek.THURSDAY));
	private static LocalDate calendarEndDate   = calendarStartDate.plusWeeks(2).minusDays(1);
	
	private LocalDate startDate;
	private LocalDate endDate;
	private LocalTime startTime;
	private String description;
	private boolean allDay;
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
		setCalendarDates(startDate);
		this.startDate   = startDate;
		this.endDate     = startDate;
		this.description = description;
		allDay = true;
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
		setCalendarDates(startDate);
		this.startDate   = startDate;
		this.startTime   = startTime;
		this.endDate     = startDate;
		this.description = description;
		allDay = false;
	}
	
	/**
	 * Sets a new end date for the event - only works for events on consecutive dates
	 * 
	 * @param endDate
	 *     the end date of the event
	 */
	public void setEndDate(LocalDate endDate) {
		if (isNextDay(endDate)) {
			setCalendarDates(endDate);
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
	 * Sets the calendar start and end dates when all dates are included
	 * 
	 * @param date
	 *     a date in the calendar
	 */
	private static void setCalendarDates(LocalDate date) {
		if (includeAllDates) {
			if (date.isBefore(calendarStartDate)) calendarStartDate = date;
			if (date.isAfter(calendarEndDate)) calendarEndDate = date;
		}
	}
	
	/**
	 * Determines if this event should be included in the output
	 * 
	 * @return
	 *     true if this event should be included in the output
	 */
	public boolean isIncluded() {
		return  (includeAllDates ||
				(startDate.compareTo(calendarStartDate) >= 0 && startDate.compareTo(calendarEndDate) <= 0) ||
				(endDate.compareTo(  calendarStartDate) >= 0 && endDate.compareTo(  calendarEndDate) <= 0));
	}
	
	/**
	 * @return
	 *     the current event in string format
	 */
	@Override
	public String toString() {
		return String.format("%-16s %-10s %s", getDateString(), getTimeString(), description);
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
				return startDate.format(DATE_FORMATTER1) + getDash() + endDate.format(DATE_FORMATTER3);
			} else {
				return startDate.format(DATE_FORMATTER1) + getDash() + endDate.format(DATE_FORMATTER2);
			}
			
		} else if (startDate.equals(lastDatePrinted)) {
			return "";
			
		} else {
			lastDatePrinted = startDate;
			return startDate.format(DATE_FORMATTER1);
		}
	}
	
	/**
	 * Resets the last date printed
	 */
	public static void resetLastDatePrinted() {
		lastDatePrinted = LocalDate.of(1970, Month.JANUARY, 1);
	}
	
	/**
	 * Gets the event time in string format
	 * 
	 * @return
	 *     the event time in string format
	 */
	public String getTimeString() {
		return allDay ? "" : startTime.format(TIME_FORMATTER1);
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
	 * Determines if all dates for calendar events should be included in the output
	 * 
	 * @param includeAllDates
	 *     the boolean parameter to determine if all dates should be included
	 */
	public static void setIncludeAllDates(boolean includeAllDates) {
		CalendarEvent.includeAllDates = includeAllDates;
	}
	
	/**
	 * Sets the start and end dates of the two week calendar
	 * 
	 * @param startDate
	 *     the start date for the two week calendar
	 */
	public static void setTwoWeekCalendarDates(LocalDate startDate) {
		calendarStartDate = startDate;
		calendarEndDate   = calendarStartDate.plusWeeks(2).minusDays(1);
	}
	
	/**
	 * Gets the calendar dates for the calendar header
	 * 
	 * @return
	 *     a string of the calendar dates
	 */
	public static String getCalendarDates() {
		return calendarStartDate.format(HEADER_DATE_FORMATTER) + getDash() + calendarEndDate.format(HEADER_DATE_FORMATTER);
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
		return (wordFormat) ? " \u2013 " : " - ";
	}
}
