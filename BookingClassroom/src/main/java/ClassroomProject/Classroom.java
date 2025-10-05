package ClassroomProject;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
// --- NEW IMPORTS FOR APACHE POI ---
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
// --- END OF NEW IMPORTS ---


public class Classroom {
    private String name;
    private String subject;
    private List<Booking> bookings;

    public Classroom(String name, String subject) {
        this.name = name;
        this.subject = subject;
        this.bookings = new ArrayList<>();
    }

    public String getName() { return name; }

    public boolean isAvailable(LocalDate date, TimeSlot timeSlot) {
        for (Booking existingBooking : bookings) {
            if (existingBooking.getDate().isEqual(date)) {
                if (existingBooking.getTimeSlot().overlapsWith(timeSlot)) {
                    return false;
                }
            }
        }
        return true;
    }

    public void addBooking(Booking newBooking) {
        this.bookings.add(newBooking);
    }

    public void displaySchedule() {
        System.out.println("--- ตารางสอนห้อง " + name + " ---");
        if (bookings.isEmpty()) {
            System.out.println("ไม่มีการจอง");
        } else {
            bookings.stream()
                    .sorted((b1, b2) -> b1.getDate().isBefore(b2.getDate()) ? -1 : 1)
                    .forEach(System.out::println);
        }
        System.out.println("--------------------------");
    }

    // --- ENTIRE NEW METHOD FOR EXPORTING TO EXCEL ---

    /**
     * Exports the current list of bookings for this classroom to an Excel .xlsx file.
     * @param filePath The path where the Excel file will be saved (e.g., "schedule.xlsx").
     */
    public void exportScheduleToExcel(String filePath) {
        // 1. Create a new Workbook. XSSFWorkbook is used for the .xlsx format.
        try (Workbook workbook = new XSSFWorkbook()) {

            // 2. Create a new Sheet in the Workbook.
            Sheet sheet = workbook.createSheet("Schedule for Room " + this.name);

            // 3. Create a Font for the header row to make it bold.
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            // 4. Create the Header Row.
            String[] columns = {"Date", "Day of Week", "Start Time", "End Time", "Booked By"};
            Row headerRow = sheet.createRow(0);

            // 5. Loop through the column titles and create header cells.
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
                cell.setCellStyle(headerCellStyle); // Apply the bold style
            }

            // 6. Fill the data rows with booking information.
            int rowNum = 1; // Start from the second row (index 1)
            // We sort the bookings by date first, just like in displaySchedule()
            List<Booking> sortedBookings = new ArrayList<>(this.bookings);
            sortedBookings.sort((b1, b2) -> b1.getDate().compareTo(b2.getDate()));

            for (Booking booking : sortedBookings) {
                Row row = sheet.createRow(rowNum++);

                // Create cells and set their values for the current booking.
                row.createCell(0).setCellValue(booking.getDate().toString()); // Date
                row.createCell(1).setCellValue(booking.getTimeSlot().getDayOfWeek().toString()); // Day
                row.createCell(2).setCellValue(booking.getTimeSlot().getStartTime().toString()); // Start Time
                row.createCell(3).setCellValue(booking.getTimeSlot().getEndTime().toString()); // End Time
                row.createCell(4).setCellValue(booking.getTeacher().getName()); // Teacher Name
            }

            // 7. Auto-size columns to fit the content.
            for (int i = 0; i < columns.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // 8. Write the workbook to an output file stream.
            // Using try-with-resources ensures the stream is closed automatically.
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }

            System.out.println("✅ Successfully exported schedule to " + filePath);

        } catch (IOException e) {
            System.err.println("❌ Failed to export schedule to Excel file.");
            e.printStackTrace();
        }
    }
}