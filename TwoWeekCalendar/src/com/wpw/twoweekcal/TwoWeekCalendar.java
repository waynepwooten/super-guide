package com.wpw.twoweekcal;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
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
 * @version 1.4
 * @since   2017-12-14
 *
 *******************************************************************************/
public class TwoWeekCalendar {
	
	// Input and output files
	public static final File CALENDAR_DATA = new File("Calendar Data.docx");		// Input file
	public static final File TWO_WEEK_CAL  = new File("Two Week Calendar.docx");	// Output file
	
	// Skipped events file
	public static final File SKIP_EVENTS      = new File("skip_events.txt");
	public static final File SKIP_IF_CONTAINS = new File("skip_if_contains.txt");
	
	// Sets of events that should be skipped
	public static final Set<String> SKIP_EVENTS_SET      = getFileContents(SKIP_EVENTS);
	public static final Set<String> SKIP_IF_CONTAINS_SET = getFileContents(SKIP_IF_CONTAINS);
	
	// Date and time patterns to determine if text read is a date or time
	public static final Pattern DATE_PATTERN = Pattern.compile("\\d{1,2}/\\d{1,2}/\\d{4}");
	public static final Pattern TIME_PATTERN = Pattern.compile("(All Day|\\d{1,2}:\\d{2}[ap])");
	
	// Date and time patterns found in the new style of calendar data
	public static final Pattern DAY_DATE_PATTERN   = Pattern.compile("[A-Z][a-z]+, ([A-Z][a-z]+) (\\d\\d?)[a-z][a-z], (\\d{4})");
	public static final Pattern TIME_EVENT_PATTERN = Pattern.compile("(All Day|(\\d{1,2}(:\\d{2})?)(am|pm)? - \\d{1,2}(:\\d{2})?(am|pm)) - (.+)");
	
	// Valid Ward Codes
	public static final Pattern WARD_CODE_PATTERN = Pattern.compile("(BP|CY|LP|CR|VV|CP|WG|GG)");
	
	// Date and time formatters to read the input date and time as a local date
	public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("M/d/yyyy");
	public static final DateTimeFormatter TIME_FORMATTER = DateTimeFormatter.ofPattern("h:mm a");
	
	// Date and time formatters to read the input date and time as a local date for new style data
	public static final DateTimeFormatter DATE_FORMATTER2 = DateTimeFormatter.ofPattern("MMMM d, yyyy");
	
	// Booleans to determine if events should be skipped
	// or the calendar should be printed to standard out
	private boolean keepAllEvents = false;
	private boolean printCalendar = false;
	
	// Boolean to determine if old or new style of calendar data will be read
	private boolean readOldStyleData = false;
	
	// A specific ward to list
	private String specificWard;
	
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
				
			} else if (arg.equals("-o")) {
				readOldStyleData = true;
				
			} else if (arg.equals("-h")) {
				showUsage();
				System.exit(0);
				
			} else if (arg.startsWith("-w")) {
				specificWard = arg.substring(2).toUpperCase();
				if (!WARD_CODE_PATTERN.matcher(specificWard).matches()) {
					System.out.println("Invalid ward pattern specified!  " + specificWard);
					System.out.println("");
					showUsage();
					System.exit(0);
				}
				
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
	 * Reads the calendar data from a Microsoft Word document.
	 * 
	 * @return
	 *     true if there were no errors reading the data
	 */
	private boolean readCalendarData() {
		if (readOldStyleData) {
			return readOldCalendarData();
		} else {
			return readNewCalendarData();
		}
	}
	
	/**
	 * Reads the calendar data from a Microsoft Word document.  The data in the Word
	 * document was copied from the church web site under Leader and Clerk Resources.
	 * 
	 * @return
	 *     true if there were no errors reading the data
	 */
	private boolean readOldCalendarData() {
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
					
					addCalendarEvent(currentDate, currentTime, text);
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
	 * Reads the calendar data from a Microsoft Word document.  The data in the Word
	 * document was copied from a print of the Agenda View of the Stake Calendar
	 * on the Church web site.
	 * 
	 * @return
	 *     true if there were no errors reading the data
	 */
	private boolean readNewCalendarData() {
		LocalDate currentDate = null;
		LocalTime currentTime = null;
		
		FileInputStream fis = null;
		XWPFDocument doc = null;
		
		try {
			fis = new FileInputStream(CALENDAR_DATA);
			doc = new XWPFDocument(OPCPackage.open(fis));
			
			for (XWPFParagraph p : doc.getParagraphs()) {
				String text = p.getText();
				
				Matcher dayDateMatcher   = DAY_DATE_PATTERN.matcher(text);
				Matcher timeEventMatcher = TIME_EVENT_PATTERN.matcher(text);
				
				if (dayDateMatcher.matches()) {
					String month = dayDateMatcher.group(1);
					String day   = dayDateMatcher.group(2);
					String year  = dayDateMatcher.group(3);
					
					String date = String.format("%s %s, %s", month, day, year);
					currentDate = LocalDate.parse(date, DATE_FORMATTER2);
					currentTime = null;
					
				} else if (timeEventMatcher.matches()) {
					if (currentDate == null) {
						System.out.println("Error:  Date must preceed any events in list!");
						return false;
					}
					
					String time      = timeEventMatcher.group(1);
					String startTime = timeEventMatcher.group(2);
					String startAmPm = timeEventMatcher.group(4);
					String endAmPm   = timeEventMatcher.group(6);
					text             = timeEventMatcher.group(7);
					
					if (time.equals("All Day")) {
						currentTime = null;
					} else {
						if (!startTime.contains(":")) startTime += ":00";
						if (startAmPm == null) startAmPm = endAmPm;
						time = startTime + " " + startAmPm.toUpperCase();
						currentTime = LocalTime.parse(time, TIME_FORMATTER);
					}
					
					addCalendarEvent(currentDate, currentTime, text);
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
	
	private void addCalendarEvent(LocalDate currentDate, LocalTime currentTime, String text) {
		if (skipEvent(text)) return;
		
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
	
	/**
	 * Adds an event to the appropriate event list
	 * 
	 * @param event
	 *     the event to add
	 */
	private void addEvent(CalendarEvent event) {
		String eventDescription = event.getDescription();
		
		if (isWardEvent(eventDescription)) {
			if (addWard(eventDescription)) {
				wardEventList.add(event);
			}
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
		String educ = eventDescription.toUpperCase();
		
		if (educ.contains("STAKE")) return false;
		if (educ.contains("SEMINARY")) return false;
		if (educ.contains("WARD CONFERENCE")) return false;
		if (educ.contains("BRANCH CONFERENCE")) return false;
		if (educ.contains("FAMILY HISTORY MARATHON")) return false;
		
		return (eventDescription.contains("BP") ||
				eventDescription.contains("Buena Park") ||
				eventDescription.contains("CY") ||
				eventDescription.contains("Cyp") ||
				eventDescription.contains("Cypress") ||
				eventDescription.contains("LP") ||
				eventDescription.contains("La Palma") ||
				eventDescription.contains("CR") ||
				eventDescription.contains("Crescent") ||
				eventDescription.contains("VV") ||
				eventDescription.contains("V V") ||
				eventDescription.contains("V. V.") ||
				eventDescription.contains("Valley View") ||
				eventDescription.contains("WG") ||
				eventDescription.contains("West Grove") ||
				eventDescription.contains("GG") ||
				eventDescription.contains("Garden Grove") ||
				eventDescription.contains("Korean"));
	}
	
	/**
	 * Determines if the ward should be added to the ward event list
	 * 
	 * @param eventDescription
	 *     the event description
	 * @return
	 *     true if no specific ward was given or if the event is for the specific ward
	 */
	private boolean addWard(String eventDescription) {
		if (specificWard == null) return true;
		
		for (String wardDescription : getWardDescriptions()) {
			if (eventDescription.contains(wardDescription)) {
				return true;
			}
		}
		
		return false;
	}
	
	/**
	 * Gets an array of ward descriptions based on the specific ward
	 * 
	 * @return
	 *     an array of ward descriptions
	 */
	private String[] getWardDescriptions() {
		if (specificWard == null) {
			return new String[] {};
		} else if (specificWard.equals("BP")) {
			return new String[] {"BP", "Buena Park Ward"};
		} else if (specificWard.equals("CY")) {
			return new String[] {"CY", "Cypress Ward"};
		} else if (specificWard.equals("LP")) {
			return new String[] {"LP", "La Palma Ward"};
		} else if (specificWard.equals("CR")) {
			return new String[] {"CR", "Crescent Ward"};
		} else if (specificWard.equals("VV")) {
			return new String[] {"VV", "V V", "V. V.", "Valley View Ward"};
		} else if (specificWard.equals("CP")) {
			return new String[] {"CP", "Cypress Park Ward"};
		} else if (specificWard.equals("WG")) {
			return new String[] {"WG", "West Grove Ward"};
		} else if (specificWard.equals("GG")) {
			return new String[] {"GG", "Garden Grove 11th Branch", "Korean"};
		} else {
			return new String[] {};
		}
	}
	
	/**
	 * Gets a set of lines in a file
	 * 
	 * @return
	 *     a set of lines in a file
	 */
	private static Set<String> getFileContents(File file) {
		Set<String> contents = new HashSet<>();
		
		BufferedReader br = null;
		String line;
		
		try {
			br = new BufferedReader(new FileReader(file));
			
			while ((line = br.readLine()) != null) {
				if (line.trim().startsWith("#")) continue;
				contents.add(line);
			}
			
		} catch (FileNotFoundException e) {
			System.out.println("File not found!  " + file);
			
		} catch (IOException e) {
			System.out.println("Error reading file!  " + file);
			e.printStackTrace();
			
		} finally {
			try {
				if (br != null) br.close();
			} catch (IOException e) {
			}
		}
		
		return contents;
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
		
		if (isSkippedEvent(text)) {
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
	 * Determines if an event is in either of the skip sets
	 * 
	 * @param text
	 *     the description of the event
	 * @return
	 *     true if the event should be skipped
	 */
	private boolean isSkippedEvent(String text) {
		for (String str : SKIP_IF_CONTAINS_SET) {
			if (text.contains(str)) return true;
		}
		
		return SKIP_EVENTS_SET.contains(text);
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
			
			CalendarEvent.resetLastDatePrinted();
			
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
		
		CalendarEvent.resetLastDatePrinted();
		
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
		System.out.println("    java -jar TwoWeekCalendar.jar [-a] [-k] [-p] [-h] [-o] [-wCODE] [start_date]");
		System.out.println("");
		System.out.println("Optional Flags:");
		System.out.println("    -a - include all dates");
		System.out.println("    -k - keep all events - do not skip");
		System.out.println("    -p - print to standard out");
		System.out.println("    -h - help - show these usage instructions");
		System.out.println("    -o - read old style calendar data");
		System.out.println("    -w - show only a specific ward");
		System.out.println("         followed by the ward code");
		System.out.println("         (BP|CY|LP|CR|VV|CP|WG|GG)");
		System.out.println("         BP=Buena Park, CY=Cypress, etc.");
		System.out.println("         (E.g. java -jar TwoWeekCalendar.jar -wBP)");
		System.out.println("");
		System.out.println("    start_date - an optional start date in the format mm/dd/yyyy");
		System.out.println("                 the default start date is the next Thursday");
		System.out.println("");
	}
}
