package fi.jonnen.whc;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Properties;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingUtilities;
import javax.swing.WindowConstants;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import fi.jonnen.whc.types.WorkHourRow;

public class WorkHourCalculatorMain {

	private static final String MESSAGE = "-------------------------------------------------------------------------";
	private static final Logger LOG = LogManager.getLogger(WorkHourCalculatorMain.class);
	private static final String DATE = "Date";
	private static final String EMPLOYEE = "Employee";
	private static final String COMMENTS = "Comments";
	private static final String PROJECT = "Project";
	private static final String TASK = "Task";
	private static final String ACTUAL_WORK = "Actual Work";

	private static final double WORK_HOURS_PER_DAY = 7.5d;

	private static final Locale LOCALE = new Locale("fi", "FI");
	private static final DateFormat DATE_FORMAT = DateFormat.getDateInstance(DateFormat.DEFAULT, LOCALE);

	private static Integer iCellDate;
	private static Integer iCellEmployee;
	private static Integer iCellComments;
	private static Integer iCellProject;
	private static Integer iCellTask;
	private static Integer iCellActualWork;

	private static Set<String> nonWorkingDates = new HashSet<>();
	private static Set<String> nonWorkingDaysOfWeek = new HashSet<>();

	public static void main(String[] args) {
		try (Stream<Path> paths = Files.walk(Paths.get("./"))) {
			loadProperties();
			List<Path> allValidFiles = paths.filter(Files::isRegularFile).filter(p -> p.toString().endsWith(".xlsx"))
					.collect(Collectors.toList());
			for (Path validFile : allValidFiles) {
				tryToReadFileAndLogReport(validFile);
			}
		} catch (Exception e) {
			LOG.error(e.getMessage(), e);
		}
		SwingUtilities.invokeLater(WorkHourCalculatorMain::createAndShowGUI);
	}

	private static void loadProperties() {
		Properties appProps = new Properties();
		try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("whc.properties")) {
			appProps.load(is);
			nonWorkingDates = Set.of(appProps.getProperty("nonWorkingDates").split(","));
			nonWorkingDaysOfWeek = Set.of(appProps.getProperty("nonWorkingDaysOfWeek").split(","));
		} catch (IOException ioe) {
			LOG.error(ioe);
		}
	}

	private static void tryToReadFileAndLogReport(Path validFile) {
		try (Workbook wb = new XSSFWorkbook(new File(validFile.toUri()))) {
			LOG.info("File: {}", validFile);
			LOG.info("Number of sheets: {}", wb.getNumberOfSheets());
			wb.forEach(sheet -> {
				LOG.info("Title of sheet => {}", sheet.getSheetName());
				sheet.removeRow(initCellIndices(sheet));
				final List<WorkHourRow> whrs = new ArrayList<>();
				final double workHourExpectedMultiplier = countDistinctEmployees(sheet);
				int iseCount = 0;
				for (Row row : sheet) {
					try {
						Cell cellDate = row.getCell(iCellDate);
						Cell cellComments = row.getCell(iCellComments);
						Cell cellProject = row.getCell(iCellProject);
						Cell cellTask = row.getCell(iCellTask);
						Cell cellActualWork = row.getCell(iCellActualWork);
						WorkHourRow whr = new WorkHourRow();
						whr.setDate(cellDate.getDateCellValue());
						whr.setComments(cellComments.getStringCellValue());
						whr.setProject(cellProject.getStringCellValue());
						whr.setTask(cellTask.getStringCellValue());
						whr.setWorkHours(cellActualWork.getNumericCellValue());
						whrs.add(whr);
					} catch (IllegalStateException ise) {
						iseCount++;
						ise.printStackTrace();
					}
				}
				logReport(whrs, workHourExpectedMultiplier, iseCount);
				resetCellIndices();
			});
		} catch (Exception e) {
			LOG.error(e.getMessage(), e);
		}
	}

	private static void logReport(final List<WorkHourRow> whrs, final double workHourExpectedMultiplier, int iseCount) {
		Collections.sort(whrs);
		Date earliestDate = whrs.get(0).getDate();
		Date latestDate = whrs.get(whrs.size() - 1).getDate();
		LOG.info("Iterating over all working dates between {} -> {}", DATE_FORMAT.format(earliestDate),
				DATE_FORMAT.format(latestDate));

		GregorianCalendar calCurrent = new GregorianCalendar(LOCALE);
		calCurrent.setTime(earliestDate);
		GregorianCalendar calEnd = new GregorianCalendar(LOCALE);
		calEnd.setTime(latestDate);
		calEnd.add(Calendar.DATE, 1);
		double workHoursDone = 0d;
		double workHoursExpected = 0d;
		final double workHoursPerDayExpected = WORK_HOURS_PER_DAY * workHourExpectedMultiplier;
		LOG.info("{}\t\t| {}\t\t| {}\t| {}", "Date", "Hours Done", "Hours Expected", "Saldo");
		LOG.info(MESSAGE);
		while (calCurrent.before(calEnd)) {
			String dayOfWeekCurrent = String.valueOf(calCurrent.get(Calendar.DAY_OF_WEEK));
			String currentDate = DATE_FORMAT.format(calCurrent.getTime());
			boolean isWorkHoursExpected = !nonWorkingDaysOfWeek.contains(dayOfWeekCurrent)
					&& !nonWorkingDates.contains(currentDate);
			List<WorkHourRow> whrsCurrent = whrs.stream().filter(whr -> calCurrent.getTime().equals(whr.getDate()))
					.collect(Collectors.toList());
			double workHoursCurrent = round(whrsCurrent.stream().mapToDouble(w -> w.getWorkHours()).sum(), 2);
			double hourBankCurrent = round(workHoursCurrent - (isWorkHoursExpected ? workHoursPerDayExpected : 0d), 2);
			if (workHoursCurrent == 0d && !isWorkHoursExpected) {
				calCurrent.add(Calendar.DATE, 1);
				continue;
			}
			LOG.info("{} {} {}\t| {}h\t\t| {}h\t\t| {}h", currentDate,
					calCurrent.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.LONG_STANDALONE, LOCALE),
					(!isWorkHoursExpected ? "!!!" : ""), workHoursCurrent,
					(isWorkHoursExpected ? workHoursPerDayExpected : 0d), hourBankCurrent);
			workHoursDone += workHoursCurrent;
			if (isWorkHoursExpected) {
				workHoursExpected += workHoursPerDayExpected;
			}
			calCurrent.add(Calendar.DATE, 1);
		}
		workHoursDone = round(workHoursDone, 2);
		workHoursExpected = round(workHoursExpected, 2);
		final double workHoursSaldo = round(workHoursDone - workHoursExpected, 2);
		LOG.info(MESSAGE);
		LOG.info("TOTAL:\t\t| {}h\t\t| {}h\t\t| {}h", workHoursDone, workHoursExpected, workHoursSaldo);
		if (iseCount > 0) {
			LOG.warn(
					"There were total of {} IllegalStateExceptions due to bad (non-data) rows. {} data rows was succesfully calculated",
					iseCount, whrs.size());
		}
		LOG.info(MESSAGE);
	}

	private static double round(double value, int places) {
		if (places < 0)
			throw new IllegalArgumentException();

		long factor = (long) Math.pow(10, places);
		value = value * factor;
		long tmp = Math.round(value);
		return (double) tmp / factor;
	}

	private static double countDistinctEmployees(Sheet sheet) {
		if (iCellEmployee == null) {
			return 1;
		}
		Set<String> employees = new HashSet<>();
		for (Row row : sheet) {
			try {
				String employee = row.getCell(iCellEmployee).getStringCellValue();
				if (!"".equals(employee)) {
					employees.add(row.getCell(iCellEmployee).getStringCellValue().toLowerCase());
				}
			} catch (IllegalStateException ise) {
				ise.printStackTrace();
			}
		}
		return !employees.isEmpty() ? employees.size() : 1;
	}

	private static void resetCellIndices() {
		iCellDate = null;
		iCellEmployee = null;
		iCellComments = null;
		iCellProject = null;
		iCellTask = null;
		iCellActualWork = null;
	}

	private static Row initCellIndices(Sheet sheet) {
		for (Row row : sheet) {
			Iterator<Cell> i = row.cellIterator();
			while (i.hasNext()) {
				Cell cell = i.next();
				if (cell.getStringCellValue().equalsIgnoreCase(DATE) && iCellDate == null) {
					iCellDate = cell.getColumnIndex();
				} else if (cell.getStringCellValue().equalsIgnoreCase(COMMENTS) && iCellComments == null) {
					iCellComments = cell.getColumnIndex();
				} else if (cell.getStringCellValue().equalsIgnoreCase(PROJECT) && iCellProject == null) {
					iCellProject = cell.getColumnIndex();
				} else if (cell.getStringCellValue().equalsIgnoreCase(TASK) && iCellTask == null) {
					iCellTask = cell.getColumnIndex();
				} else if (cell.getStringCellValue().equalsIgnoreCase(ACTUAL_WORK) && iCellActualWork == null) {
					iCellActualWork = cell.getColumnIndex();
				}

				// Used only for distinct count, and may not be present
				if (cell.getStringCellValue().equalsIgnoreCase(EMPLOYEE)) {
					iCellEmployee = cell.getColumnIndex();
				}
			}
			if (iCellDate != null && iCellComments != null && iCellProject != null && iCellTask != null
					&& iCellActualWork != null) {
				return row;
			}
		}
		LOG.warn("Could not initCellIndices! Make sure that excel file contains proper header row!");
		return null;
	}

	private static void createAndShowGUI() {
		JTextArea textArea = new JTextArea(5, 20);
		textArea.setEditable(false);
		try (Stream<String> lines = Files.lines(Paths.get("./log/log4j2.log"), StandardCharsets.UTF_8)) {
			lines.forEach(s -> textArea.append(s + "\n"));
		} catch (IOException ioe) {
			LOG.error(ioe.getMessage(), ioe);
		}
		JScrollPane areaScrollPane = new JScrollPane(textArea);
		areaScrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
		areaScrollPane.setPreferredSize(new Dimension(800, 600));
		JFrame frame = new JFrame("Work Hour Calculator Log Viewer");
		frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		frame.getContentPane().add(areaScrollPane);

		frame.pack();
		frame.setVisible(true);
	}

}
