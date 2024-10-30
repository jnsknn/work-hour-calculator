package fi.jonnen.whc;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
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

import fi.jonnen.whc.types.WorkHourRow;

public class WorkHourCalculatorMain {

	private static final Logger LOG = LogManager.getLogger(WorkHourCalculatorMain.class);
	private static final String DATE = "Date";
	private static final String COMMENTS = "Comments";
	private static final String PROJECT = "Project";
	private static final String TASK = "Task";
	private static final String ACTUAL_WORK = "Actual Work";

	private static final double WORK_HOURS_PER_DAY = 7.5d;

	private static final Locale LOCALE = new Locale("fi", "FI");
	private static final DateFormat DATE_FORMAT = DateFormat.getDateInstance(DateFormat.DEFAULT, LOCALE);

	private static Integer iCellDate;
	private static Integer iCellComments;
	private static Integer iCellProject;
	private static Integer iCellTask;
	private static Integer iCellActualWork;
	
	private static Set<String> nonWorkingDates = new HashSet<String>();
	private static Set<String> nonWorkingDaysOfWeek = new HashSet<String>();

	public static void main(String[] args) {
		loadProperties();
		try (Stream<Path> paths = Files.walk(Paths.get("./"))) {
			List<Path> allValidFiles = paths.filter(Files::isRegularFile).filter(p -> p.toString().endsWith(".xlsx"))
					.collect(Collectors.toList());
			for (Path validFile : allValidFiles) {
				tryToReadFile(validFile);
			}
		} catch (IOException ioe) {
			LOG.error(ioe.getMessage(), ioe);
		}
		SwingUtilities.invokeLater(WorkHourCalculatorMain::createAndShowGUI);
	}

	private static void loadProperties() {
		String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
		String appConfigPath = rootPath + "app.properties";
		Properties appProps = new Properties();
		try (FileInputStream fis = new FileInputStream(appConfigPath)) {
			appProps.load(fis);
			nonWorkingDates = Set.of(appProps.getProperty("nonWorkingDates").split(","));
			nonWorkingDaysOfWeek = Set.of(appProps.getProperty("nonWorkingDaysOfWeek").split(","));
		} catch (IOException ioe) {
			LOG.error(ioe);
		} 
	}

	private static void tryToReadFile(Path validFile) {
		try (Workbook wb = WorkbookFactory.create(new File(validFile.toUri()))) {
			LOG.info("File: {}", validFile);
			LOG.info("Number of sheets: {}", wb.getNumberOfSheets());
			wb.forEach(sheet -> {
				LOG.info("Title of sheet => {}", sheet.getSheetName());
				sheet.removeRow(initCellIndices(sheet));
				final List<WorkHourRow> whrs = new ArrayList<>();
				for (Row row : sheet) {
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
				}
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
				LOG.info("{}\t\t| {}\t\t| {}", "Date", "Hours", "Saldo");
				LOG.info("---------------------------------------------------------");
				while (calCurrent.before(calEnd)) {
					List<WorkHourRow> whrsCurrent = whrs.stream()
							.filter(whr -> calCurrent.getTime().equals(whr.getDate())).collect(Collectors.toList());
					double workHoursCurrent = whrsCurrent.stream().mapToDouble(w -> w.getWorkHours()).sum();
					double hourBankCurrent = workHoursCurrent - WORK_HOURS_PER_DAY;
					String dayOfWeekCurrent = String.valueOf(calCurrent.get(Calendar.DAY_OF_WEEK));
					String currentDate = DATE_FORMAT.format(calCurrent.getTime());
					if(workHoursCurrent == 0d && (nonWorkingDaysOfWeek.contains(dayOfWeekCurrent) || nonWorkingDates.contains(currentDate))) {
						calCurrent.add(Calendar.DATE, 1);
						continue;
					}
					LOG.info("{} {}\t| {}h/{}h\t\t| {}h", currentDate,
							calCurrent.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.LONG_STANDALONE, LOCALE),
							workHoursCurrent, WORK_HOURS_PER_DAY, hourBankCurrent);
					workHoursDone += workHoursCurrent;
					workHoursExpected += WORK_HOURS_PER_DAY;
					calCurrent.add(Calendar.DATE, 1);
				}
				LOG.info("---------------------------------------------------------");
				LOG.info("TOTAL:\t\t| {}h/{}h\t| {}h", workHoursDone, workHoursExpected,
						(workHoursDone - workHoursExpected));
			});
		} catch (Exception e) {
			LOG.error(e.getMessage(), e);
		}
	}

	private static Row initCellIndices(Sheet sheet) {
		for (Row row : sheet) {
			Iterator<Cell> i = row.cellIterator();
			while (i.hasNext()) {
				Cell cell = i.next();
				if (cell.getStringCellValue().equals(DATE) && iCellDate == null) {
					iCellDate = cell.getColumnIndex();
				} else if (cell.getStringCellValue().equals(COMMENTS) && iCellComments == null) {
					iCellComments = cell.getColumnIndex();
				} else if (cell.getStringCellValue().equals(PROJECT) && iCellProject == null) {
					iCellProject = cell.getColumnIndex();
				} else if (cell.getStringCellValue().equals(TASK) && iCellTask == null) {
					iCellTask = cell.getColumnIndex();
				} else if (cell.getStringCellValue().equals(ACTUAL_WORK) && iCellActualWork == null) {
					iCellActualWork = cell.getColumnIndex();
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
