import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class Subject {
    private String id;
    private String name;

    private static final String SUBJECTS_XLSX_FILE_PATH = "./files/excel_files/subject.xlsx";

    public Subject(String id, String name) {
        this.id = id;
        this.name = name;
    }

    public static String getNameById(String id) throws IOException, InvalidFormatException {
        ArrayList<Subject> subjects = new ArrayList<>();

        Workbook workbook = WorkbookFactory.create(new File(SUBJECTS_XLSX_FILE_PATH));

        // Получение листа по индексу (первый лист)
        Sheet sheet = workbook.getSheetAt(0);

        // Создание DataFormatter для форматирования и получения значения ячейки как String
        DataFormatter dataFormatter = new DataFormatter();

        // цикл for-each для итерации по строкам и столбцам
        for (Row row: sheet) {
            if (row.getRowNum() != 0) {
                String gotId = dataFormatter.formatCellValue(row.getCell(0));
                String gotName = dataFormatter.formatCellValue(row.getCell(1));

                Subject subject = new Subject(gotId, gotName);
                subjects.add(subject);
            }
        }

        // Закрытие workbook
        workbook.close();

        String neededName = "";
        for (Subject subject : subjects) {
            if (id.equals(subject.getId())) {
                neededName = subject.getName();
            }
        }

        return neededName;
    }

    public String getId() {
        return id;
    }

    public String getName() {
        return name;
    }
}
