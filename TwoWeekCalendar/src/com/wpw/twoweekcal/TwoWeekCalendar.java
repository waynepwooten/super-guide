package com.wpw.twoweekcal;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabs;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;

/*******************************************************************************
 * Creates a two week calendar for publishing to the Cypress California Stake
 * 
 * This class reads calendar data that has been copied from the church web site
 * under Leader and Clerk Resources and pasted into a Microsoft Word document.
 * It reads the data from the MS Word document and creates a new MS Word
 * document with the desired data formatted in the required format.
 * 
 * @author  Wayne Wooten
 * @version 1.1
 * @since   2017-12-14
 *
 *******************************************************************************/
public class TwoWeekCalendar {
	
	// Input and output files
	public static final File CALENDAR_DATA = new File("Calendar Data.docx");		// Input file
	public static final File TWO_WEEK_CAL  = new File("Two Week Calendar.docx");	// Output file
	
	// Date and time patterns to determine if text read is a date or time
	public static final Pattern DATE_PATTERN = Pattern.compile("\\d{1,2}/\\d{1,2}/\\d{4}");
	public static final Pattern TIME_PATTERN = Pattern.compile("(All Day|\\d{1,2}:\\d{2}[ap])");
	
	// Date and time formatters to read the input date and time as a local date
	public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("M/d/yyyy");
	public static final DateTimeFormatter TIME_FORMATTER = DateTimeFormatter.ofPattern("h:mm a");
	
	// A list of events that should be skipped
	public static final List<String> SKIPPED_EVENT_LIST = getSkippedEventsList();
	
	// Booleans to determine if events should be skipped
	// or the calendar should be printed to standard out
	boolean keepAllEvents = false;
	boolean printCalendar = false;
	
	// Maps of events that are all day events
	// and events that have been skipped
	private Map<String, CalendarEvent> allDayEventMap  = new HashMap<>();
	private Map<String, Integer>       skippedEventMap = new TreeMap<>();
	
	// Lists of stake and ward events
	private List<CalendarEvent> stakeEventList = new ArrayList<>();
	private List<CalendarEvent> wardEventList  = new ArrayList<>();
	
	
	/**
	 * Main method for the two week calendar application
	 * 
	 * @param args
	 *     the arguments passed to the application
	 */
	public static void main(String[] args) {
		TwoWeekCalendar twoWeekCalendar = new TwoWeekCalendar();
		twoWeekCalendar.parseArgs(args);
		twoWeekCalendar.run();
	}
	
	/**
	 * Parses the arguments passed to the application
	 * and sets the appropriate settings
	 * 
	 * @param args
	 *     the arguments passed to the application
	 */
	private void parseArgs(String[] args) {
		for (String arg : args) {
			if (arg.equals("-a" )) {
				CalendarEvent.setIncludeAllDates(true);
				
			} else if (arg.equals("-k")) {
				keepAllEvents = true;
				
			} else if (arg.equals("-p")) {
				CalendarEvent.setWordFormat(false);
				printCalendar = true;
				
			} else if (arg.equals("-h")) {
				showUsage();
				System.exit(0);
				
			} else {
				if (DATE_PATTERN.matcher(arg).matches()) {
					LocalDate startDate = LocalDate.parse(arg, DATE_FORMATTER);
					CalendarEvent.setTwoWeekCalendarDates(startDate);
				} else {
					System.out.println("Invalid argument passed!  " + arg);
					System.out.println("");
					showUsage();
					System.exit(1);
				}
			}
		}
	}
	
	/**
	 * Calls the methods to read the calendar data and create the two week calendar
	 */
	private void run() {
		if (readCalendarData()) {
			if (printCalendar) {
				printCalendar();
			} else {
				writeCalendar();
			}
		}
	}
	
	/**
	 * Reads the calendar data from a Microsoft Word document.  The data in the Word
	 * document was copied from the church web site under Leader and Clerk Resources.
	 * 
	 * @return
	 *     true if there were no errors reading the data
	 */
	private boolean readCalendarData() {
		LocalDate currentDate = null;
		LocalTime currentTime = null;
		
		FileInputStream fis = null;
		XWPFDocument doc = null;
		
		try {
			fis = new FileInputStream(CALENDAR_DATA);
			doc = new XWPFDocument(OPCPackage.open(fis));
			
			for (XWPFParagraph p : doc.getParagraphs()) {
				String text = p.getText();
				
				if (DATE_PATTERN.matcher(text).matches()) {
					currentDate = LocalDate.parse(text, DATE_FORMATTER);
					currentTime = null;
					
				} else if (TIME_PATTERN.matcher(text).matches()) {
					if (text.equals("All Day")) {
						currentTime = null;
					} else {
						text = text.substring(0, text.length()-1) + (text.endsWith("a") ? " AM" : " PM");
						currentTime = LocalTime.parse(text, TIME_FORMATTER);
					}
					
				} else {
					if (currentDate == null) {
						System.out.println("Error:  Date must be first item in list!");
						return false;
					}
					
					if (skipEvent(text)) continue;
					
					CalendarEvent event = null;
					
					if (currentTime == null) {
						if (allDayEventMap.containsKey(text)) {
							event = allDayEventMap.get(text);
							
							if (event.isNextDay(currentDate)) {
								event.setEndDate(currentDate);
							} else {
								event = new CalendarEvent(currentDate, text);
								addEvent(event);
								allDayEventMap.put(text, event);
							}
							
						} else {
							event = new CalendarEvent(currentDate, text);
							addEvent(event);
							allDayEventMap.put(text, event);
						}
						
					} else {
						event = new CalendarEvent(currentDate, currentTime, text);
						addEvent(event);
					}
				}
			}
			
			return true;
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return false;
		} catch (InvalidFormatException e) {
			e.printStackTrace();
			return false;
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		} finally {
			try {
				if (doc != null) doc.close();
				if (fis != null) fis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	/**
	 * Adds an event to the appropriate event list
	 * 
	 * @param event
	 *     the event to add
	 */
	private void addEvent(CalendarEvent event) {
		if (isWardEvent(event.getDescription())) {
			wardEventList.add(event);
		} else {
			stakeEventList.add(event);
		}
	}
	
	/**
	 * Determines if the event is a ward event by looking for 
	 * specific character patterns in the event description
	 * 
	 * @param eventDescription
	 *     the event description
	 * @return
	 *     true if the description contains ward specific character patterns
	 */
	private boolean isWardEvent(String eventDescription) {
		if (eventDescription.contains("WARD CONFERENCE")) return false;
		
		return (eventDescription.contains("BP") ||
				eventDescription.contains("Buena Park Ward") ||
				eventDescription.contains("CY") ||
				eventDescription.contains("Cypress Ward") ||
				eventDescription.contains("LP") ||
				eventDescription.contains("La Palma Ward") ||
				eventDescription.contains("CR") ||
				eventDescription.contains("Crescent Ward") ||
				eventDescription.contains("VV") ||
				eventDescription.contains("V V") ||
				eventDescription.contains("V. V.") ||
				eventDescription.contains("Valley View Ward") ||
				eventDescription.contains("CP") ||
				eventDescription.contains("Cypress Park Ward") ||
				eventDescription.contains("WG") ||
				eventDescription.contains("West Grove Ward") ||
				eventDescription.contains("GG") ||
				eventDescription.contains("Garden Grove 11th Branch") ||
				eventDescription.contains("Korean"));
	}
	
	/**
	 * Gets a list of events that should be skipped
	 * 
	 * @return
	 *     a list of events that should be skipped
	 */
	private static List<String> getSkippedEventsList() {
		List<String> skippedEvents = new ArrayList<>();
		skippedEvents.add("FHE");
		skippedEvents.add("Stake Employment Center Hours");
		skippedEvents.add("Stake Bishop's Baptisms");
		return skippedEvents;
	}
	
	/**
	 * Determines if an event should be skipped and
	 * adds skipped events to the skipped event map  
	 * 
	 * @param text
	 *     the description of the event
	 * @return
	 *     true if the event should be skipped
	 */
	private boolean skipEvent(String text) {
		if (keepAllEvents) return false;
		
		if (text.contains("ARP") || text.contains("PASG") || SKIPPED_EVENT_LIST.contains(text)) {
			int count = 0;
			
			if (skippedEventMap.containsKey(text)) {
				count = skippedEventMap.get(text);
			}
			
			skippedEventMap.put(text, ++count);
			return true;
			
		} else {
			return text.trim().isEmpty();
		}
	}
	
	/**
	 * Writes the formatted calendar information to a Microsoft Word document
	 */
	private void writeCalendar() {
		
		XWPFDocument doc = null;
		FileOutputStream fos = null;
		XWPFParagraph p;
		XWPFRun run;
		
		try {
			doc = new XWPFDocument();
			
			p = doc.createParagraph();
			setParagraph(p);
			run = p.createRun();
			run.setFontFamily("Times New Roman");
			run.setBold(true);
			run.addTab();
			run.addTab();
			run.setText(CalendarEvent.getCalendarDates());
			
			p = doc.createParagraph();
			setParagraph(p);
			run = p.createRun();
			run.setFontFamily("Times New Roman");
			run.setUnderline(UnderlinePatterns.SINGLE);
			run.setBold(true);
			run.setText("Stake-wide");
			
			for (CalendarEvent event : stakeEventList) {
				if (event.isIncluded()) {
					p = doc.createParagraph();
					setParagraph(p);
					run = p.createRun();
					run.setFontFamily("Times New Roman");
					run.setText(event.getDateString());
					run.addTab();
					run.setText(event.getTimeString());
					run.addTab();
					run.setText(event.getDescription());
				}
			}
			
			p = doc.createParagraph();
			setParagraph(p);
			run = p.createRun();
			run.setFontFamily("Times New Roman");
			run.setText("");
			
			p = doc.createParagraph();
			setParagraph(p);
			run = p.createRun();
			run.setFontFamily("Times New Roman");
			run.setUnderline(UnderlinePatterns.SINGLE);
			run.setBold(true);
			run.setText("Ward Specific");
			
			for (CalendarEvent event : wardEventList) {
				if (event.isIncluded()) {
					p = doc.createParagraph();
					setParagraph(p);
					run = p.createRun();
					run.setFontFamily("Times New Roman");
					run.setText(event.getDateString());
					run.addTab();
					run.setText(event.getTimeString());
					run.addTab();
					run.setText(event.getDescription());
				}
			}
			
			fos = new FileOutputStream(TWO_WEEK_CAL);
			doc.write(fos);
			fos.close();
			
			printSkippedEvents();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (doc != null) doc.close();
				if (fos != null) fos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	/**
	 * Sets the formatting of a paragraph in the MS Word document
	 * 
	 * @param p
	 *     the MS Word paragraph
	 */
	private void setParagraph(XWPFParagraph p) {
		p.setIndentationHanging(3600);
		p.setIndentationLeft(3600);
		p.setIndentFromLeft(3600);
		p.setSpacingAfter(0);
		p.setSpacingBetween(1.0);
		setTabStop(p, STTabJc.LEFT, new BigInteger("2160"));
		setTabStop(p, STTabJc.LEFT, new BigInteger("3600"));
	}
	
	/**
	 * Sets a tab stop in a paragraph in the MS Word document
	 * 
	 * @param p
	 *     the MS Word paragraph
	 * @param tabType
	 *     the tab type (Left, Center, Right, etc.)
	 * @param pos
	 *     the position of the tab (1440 = 1 inch)
	 */
	private void setTabStop(XWPFParagraph p, STTabJc.Enum tabType, BigInteger pos) {
		CTP ctp = p.getCTP();
		
		CTPPr ppr = ctp.getPPr();
		if (ppr == null) {
			ppr = ctp.addNewPPr();
		}
		
		CTTabs tabs = ppr.getTabs();
		if (tabs == null) {
			tabs = ppr.addNewTabs();
		}
		
		CTTabStop tabStop = tabs.addNewTab();
		tabStop.setVal(tabType);
		tabStop.setPos(pos);
	}
	
	/**
	 * Prints the calendar to standard out
	 */
	private void printCalendar() {
		System.out.println("");
		System.out.printf("%28s%s%n", "", CalendarEvent.getCalendarDates());
		System.out.println("Stake-wide");
		for (CalendarEvent event : stakeEventList) {
			if (event.isIncluded()) {
				System.out.println(event);
			}
		}
		System.out.println("");
		System.out.println("Ward Specific");
		for (CalendarEvent event : wardEventList) {
			if (event.isIncluded()) {
				System.out.println(event);
			}
		}
		printSkippedEvents();
	}
	
	/**
	 * Prints a list of skipped events
	 */
	private void printSkippedEvents() {
		if (!keepAllEvents) {
			System.out.println("");
			System.out.println("SKIPPED EVENTS");
			for (String text : skippedEventMap.keySet()) {
				System.out.println(text + " (" + skippedEventMap.get(text) + ")");
			}
			System.out.println("");
		}
	}
	
	/**
	 * Shows the application usage information
	 */
	private void showUsage() {
		System.out.println("");
		System.out.println("This application creates a two week calendar for publishing to the");
		System.out.println("Cypress California Stake.");
		System.out.println("");
		System.out.println("Usage:");
		System.out.println("    java -jar TwoWeekCalendar.jar [-a] [-k] [-p] [-h] [start_date]");
		System.out.println("");
		System.out.println("Optional Flags:");
		System.out.println("    -a - include all dates");
		System.out.println("    -k - keep all events - do not skip");
		System.out.println("    -p - print to standard out");
		System.out.println("    -h - help - show these usage instructions");
		System.out.println("");
		System.out.println("    start_date - an optional start date in the format mm/dd/yyyy");
		System.out.println("                 the default start date is the next Thursday");
		System.out.println("");
	}
}
