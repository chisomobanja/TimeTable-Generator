

// Now the enhanced Java application:
import java.io.FileOutputStream;
import java.time.DayOfWeek;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Subject implements Comparable<Subject> {
    private String name;
    private int duration; // in minutes
    private DayOfWeek day;
    private LocalTime startTime;
    private int id;

    public Subject(int id, String name, int duration, DayOfWeek day, LocalTime startTime) {
        this.id = id;
        this.name = name;
        this.duration = duration;
        this.day = day;
        this.startTime = startTime;
    }

    // Getters and setters
    public int getId() { return id; }
    public String getName() { return name; }
    public void setName(String name) { this.name = name; }
    public int getDuration() { return duration; }
    public void setDuration(int duration) { this.duration = duration; }
    public DayOfWeek getDay() { return day; }
    public void setDay(DayOfWeek day) { this.day = day; }
    public LocalTime getStartTime() { return startTime; }
    public void setStartTime(LocalTime startTime) { this.startTime = startTime; }

    public LocalTime getEndTime() {
        return startTime.plusMinutes(duration);
    }

    @Override
    public int compareTo(Subject other) {
        int dayCompare = this.day.compareTo(other.day);
        if (dayCompare != 0) return dayCompare;
        return this.startTime.compareTo(other.startTime);
    }
}

public class TimetableApplication {
    private static ArrayList<Subject> subjects = new ArrayList<>();
    private static Scanner scanner = new Scanner(System.in);
    private static int nextId = 1;

    public static void main(String[] args) {
        boolean running = true;

        while (running) {
            System.out.println("\n=== Timetable Application ===");
            System.out.println("1. Add Subject");
            System.out.println("2. Edit Subject");
            System.out.println("3. Delete Subject");
            System.out.println("4. View All Subjects");
            System.out.println("5. Check Time Conflicts");
            System.out.println("6. Sort Subjects");
            System.out.println("7. Export to Excel");
            System.out.println("8. Exit");
            System.out.print("Choose an option: ");

            int choice = scanner.nextInt();
            scanner.nextLine(); // consume newline

            switch (choice) {
                case 1: addSubject(); break;
                case 2: editSubject(); break;
                case 3: deleteSubject(); break;
                case 4: viewSubjects(); break;
                case 5: checkConflicts(); break;
                case 6: sortSubjects(); break;
                case 7: exportToExcel(); break;
                case 8: running = false; break;
                default: System.out.println("Invalid option!");
            }
        }
        scanner.close();
    }

    private static void addSubject() {
        System.out.print("Enter subject name: ");
        String name = scanner.nextLine();

        System.out.print("Enter duration (in minutes): ");
        int duration = scanner.nextInt();
        scanner.nextLine();

        System.out.print("Enter day (1-7, where 1 is Monday): ");
        int dayNum = scanner.nextInt();
        scanner.nextLine();
        DayOfWeek day = DayOfWeek.of(dayNum);

        System.out.print("Enter start time (HH:mm): ");
        String timeStr = scanner.nextLine();
        LocalTime startTime = LocalTime.parse(timeStr);

        Subject newSubject = new Subject(nextId++, name, duration, day, startTime);
        if (hasTimeConflict(newSubject)) {
            System.out.println("Warning: This subject has time conflicts with existing subjects!");
            System.out.print("Add anyway? (y/n): ");
            if (scanner.nextLine().toLowerCase().startsWith("n")) {
                return;
            }
        }
        subjects.add(newSubject);
        System.out.println("Subject added successfully!");
    }

    private static void editSubject() {
        viewSubjects();
        System.out.print("Enter subject ID to edit: ");
        int id = scanner.nextInt();
        scanner.nextLine();

        Subject subject = subjects.stream()
                .filter(s -> s.getId() == id)
                .findFirst()
                .orElse(null);

        if (subject == null) {
            System.out.println("Subject not found!");
            return;
        }

        System.out.println("Leave blank to keep current value");
        
        System.out.print("New name [" + subject.getName() + "]: ");
        String name = scanner.nextLine();
        if (!name.isEmpty()) subject.setName(name);

        System.out.print("New duration in minutes [" + subject.getDuration() + "]: ");
        String durationStr = scanner.nextLine();
        if (!durationStr.isEmpty()) subject.setDuration(Integer.parseInt(durationStr));

        System.out.print("New day (1-7, where 1 is Monday) [" + subject.getDay().getValue() + "]: ");
        String dayStr = scanner.nextLine();
        if (!dayStr.isEmpty()) subject.setDay(DayOfWeek.of(Integer.parseInt(dayStr)));

        System.out.print("New start time (HH:mm) [" + subject.getStartTime() + "]: ");
        String timeStr = scanner.nextLine();
        if (!timeStr.isEmpty()) subject.setStartTime(LocalTime.parse(timeStr));

        if (hasTimeConflict(subject)) {
            System.out.println("Warning: This change creates time conflicts!");
        }
        
        System.out.println("Subject updated successfully!");
    }

    private static void deleteSubject() {
        viewSubjects();
        System.out.print("Enter subject ID to delete: ");
        int id = scanner.nextInt();
        scanner.nextLine();

        subjects.removeIf(s -> s.getId() == id);
        System.out.println("Subject deleted (if found)!");
    }

    private static void viewSubjects() {
        if (subjects.isEmpty()) {
            System.out.println("No subjects found!");
            return;
        }

        System.out.println("\nCurrent Subjects:");
        System.out.printf("%-4s %-20s %-10s %-15s %-15s %-15s%n", 
                "ID", "Name", "Day", "Start Time", "End Time", "Duration");
        System.out.println("-".repeat(80));

        for (Subject subject : subjects) {
            System.out.printf("%-4d %-20s %-10s %-15s %-15s %-15d%n",
                    subject.getId(),
                    subject.getName(),
                    subject.getDay(),
                    subject.getStartTime(),
                    subject.getEndTime(),
                    subject.getDuration());
        }
    }

    private static boolean hasTimeConflict(Subject subject) {
        return subjects.stream()
                .filter(s -> s.getId() != subject.getId()) // Skip self when checking conflicts
                .filter(s -> s.getDay() == subject.getDay())
                .anyMatch(s -> {
                    LocalTime newStart = subject.getStartTime();
                    LocalTime newEnd = subject.getEndTime();
                    LocalTime existingStart = s.getStartTime();
                    LocalTime existingEnd = s.getEndTime();

                    return !(newEnd.isBefore(existingStart) || newStart.isAfter(existingEnd));
                });
    }

    private static void checkConflicts() {
        System.out.println("\nChecking for time conflicts...");
        boolean foundConflicts = false;

        for (int i = 0; i < subjects.size(); i++) {
            for (int j = i + 1; j < subjects.size(); j++) {
                Subject s1 = subjects.get(i);
                Subject s2 = subjects.get(j);

                if (s1.getDay() == s2.getDay()) {
                    LocalTime s1Start = s1.getStartTime();
                    LocalTime s1End = s1.getEndTime();
                    LocalTime s2Start = s2.getStartTime();
                    LocalTime s2End = s2.getEndTime();

                    if (!(s1End.isBefore(s2Start) || s1Start.isAfter(s2End))) {
                        foundConflicts = true;
                        System.out.printf("Conflict found between:%n" +
                                "  %s (%s %s-%s)%n" +
                                "  %s (%s %s-%s)%n",
                                s1.getName(), s1.getDay(), s1Start, s1End,
                                s2.getName(), s2.getDay(), s2Start, s2End);
                    }
                }
            }
        }

        if (!foundConflicts) {
            System.out.println("No time conflicts found!");
        }
    }

    private static void sortSubjects() {
        System.out.println("\nSort by:");
        System.out.println("1. Day and Time");
        System.out.println("2. Subject Name");
        System.out.println("3. Duration");
        System.out.print("Choose sorting method: ");

        int choice = scanner.nextInt();
        scanner.nextLine();

        switch (choice) {
            case 1:
                Collections.sort(subjects);
                break;
            case 2:
                subjects.sort(Comparator.comparing(Subject::getName));
                break;
            case 3:
                subjects.sort(Comparator.comparing(Subject::getDuration));
                break;
            default:
                System.out.println("Invalid sorting option!");
                return;
        }

        System.out.println("Subjects sorted successfully!");
        viewSubjects();
    }

    private static void exportToExcel() {
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create weekly view sheet
            Sheet weeklySheet = workbook.createSheet("Weekly View");
            createWeeklyView(weeklySheet);

            // Create list view sheet
            Sheet listSheet = workbook.createSheet("List View");
            createListView(listSheet);

            // Write to file
            try (FileOutputStream fileOut = new FileOutputStream("Timetable.xlsx")) {
                workbook.write(fileOut);
                System.out.println("Excel file has been created successfully!");
            }

        } catch (Exception e) {
            System.out.println("Error creating Excel file: " + e.getMessage());
        }
    }

    private static void createWeeklyView(Sheet sheet) {
        // Create header row with days
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Time");
        for (int i = 1; i <= 7; i++) {
            headerRow.createCell(i).setCellValue(DayOfWeek.of(i).toString());
        }

        // Create time slots (30-minute intervals from 8:00 to 20:00)
        LocalTime startTime = LocalTime.of(8, 0);
        LocalTime endTime = LocalTime.of(20, 0);
        int rowNum = 1;
        
        while (!startTime.isAfter(endTime)) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(startTime.toString());
            startTime = startTime.plusMinutes(30);
        }

        // Fill in subjects
        CellStyle subjectStyle = workbook.createCellStyle();
        subjectStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        subjectStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (Subject subject : subjects) {
            int dayColumn = subject.getDay().getValue();
            LocalTime subjectStart = subject.getStartTime();
            int startRow = (subjectStart.getHour() - 8) * 2 + 1;
            if (subjectStart.getMinute() >= 30) startRow++;

            Row row = sheet.getRow(startRow);
            if (row == null) continue;

            Cell cell = row.createCell(dayColumn);
            cell.setCellValue(subject.getName());
            cell.setCellStyle(subjectStyle);
        }

        // Autosize columns
        for (int i = 0; i <= 7; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private static void createListView(Sheet sheet) {
        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID");
        headerRow.createCell(1).setCellValue("Subject");
        headerRow.createCell(2).setCellValue("Day");
        headerRow.createCell(3).setCellValue("Start Time");
        headerRow.createCell(4).setCellValue("End Time");
        headerRow.createCell(5).setCellValue("Duration (minutes)");

        // Add data rows
        int rowNum = 1;
        for (Subject subject : subjects) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(subject.getId());
            row.createCell(1).setCellValue(subject.getName());
            row.createCell(2).setCellValue(subject.getDay().toString());
            row.createCell(3).setCellValue(subject.getStartTime().toString());
            row.createCell(4).setCellValue(subject.getEndTime().toString());
            row.createCell(5).setCellValue(subject.getDuration());
        }

        // Autosize columns
        for (int i = 0; i <= 5; i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
