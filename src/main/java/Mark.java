import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class Mark {
    private String id;
    private String subjectId;
    private String value;
    private String studentId;

    private static final String MARK_XLSX_FILE_PATH = "./files/excel_files/mark.xlsx";

    public Mark(String id, String subjectId, String value, String studentId) {
        this.id = id;
        this.subjectId = subjectId;
        this.value = value;
        this.studentId = studentId;
    }

    public static ArrayList<Mark> readMarksOfStudent(String studentId) throws IOException, InvalidFormatException {
        ArrayList<Mark> marks = new ArrayList<>();

        Workbook workbook = WorkbookFactory.create(new File(MARK_XLSX_FILE_PATH));

        // Получение листа по индексу (первый лист)
        Sheet sheet = workbook.getSheetAt(0);

        // Создание DataFormatter для форматирования и получения значения ячейки как String
        DataFormatter dataFormatter = new DataFormatter();

        // цикл for-each для итерации по строкам и столбцам
        for (Row row: sheet) {
            if (row.getRowNum() != 0) {
                String gotId = dataFormatter.formatCellValue(row.getCell(0));
                String gotSubjectIdId = dataFormatter.formatCellValue(row.getCell(1));
                String gotValue = dataFormatter.formatCellValue(row.getCell(2));
                String gotStudentId = dataFormatter.formatCellValue(row.getCell(3));

                if (studentId.equals(gotStudentId)) {
                    Mark mark = new Mark(gotId, gotSubjectIdId, gotValue, gotStudentId);
                    marks.add(mark);
                }
            }
        }

        // Закрытие workbook
        workbook.close();

        return marks;
    }

    public String getId() {
        return id;
    }

    public String getSubjectId() {
        return subjectId;
    }

    public String getValue() {
        return value;
    }

    public String getStudentId() {
        return studentId;
    }
}
