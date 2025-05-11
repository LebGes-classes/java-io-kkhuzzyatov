import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class Schedule {
    private String id;
    private String educatorId;
    private String subjectId;
    private String day;
    private String startTime;
    private String group;

    private static final String SCHEDULE_XLSX_FILE_PATH = "./files/excel_files/schedule.xlsx";

    public Schedule(String id, String educatorId, String subjectId, String day, String startTime, String group) {
        this.id = id;
        this.educatorId = educatorId;
        this.subjectId = subjectId;
        this.day = day;
        this.startTime = startTime;
        this.group = group;
    }

    public static ArrayList<Schedule> readSchedulesOfEducator(String educatorId) throws IOException, InvalidFormatException {
        ArrayList<Schedule> schedules = new ArrayList<>();

        Workbook workbook = WorkbookFactory.create(new File(SCHEDULE_XLSX_FILE_PATH));

        // Получение листа по индексу (первый лист)
        Sheet sheet = workbook.getSheetAt(0);

        // Создание DataFormatter для форматирования и получения значения ячейки как String
        DataFormatter dataFormatter = new DataFormatter();

        // цикл for-each для итерации по строкам и столбцам
        for (Row row: sheet) {
            if (row.getRowNum() != 0) {
                String gotId = dataFormatter.formatCellValue(row.getCell(0));
                String gotEducatorId = dataFormatter.formatCellValue(row.getCell(1));
                String gotSubjectId = dataFormatter.formatCellValue(row.getCell(2));
                String gotDay = dataFormatter.formatCellValue(row.getCell(3));
                String gotStartTime = dataFormatter.formatCellValue(row.getCell(4));
                String gotGroup = dataFormatter.formatCellValue(row.getCell(5));

                if (educatorId.equals(gotEducatorId)) {
                    Schedule schedule = new Schedule(gotId, gotEducatorId, gotSubjectId, gotDay, gotStartTime, gotGroup);
                    schedules.add(schedule);
                }
            }
        }

        // Закрытие workbook
        workbook.close();

        return schedules;
    }

    public String getId() {
        return id;
    }

    public String getSubjectId() {
        return subjectId;
    }

    public String getDay() {
        return day;
    }

    public String getStartTime() {
        return startTime;
    }

    public String getGroup() {
        return group;
    }
}
