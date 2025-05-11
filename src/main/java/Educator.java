import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.ArrayList;

public class Educator {
    private String id;
    private String name;
    private ArrayList<Schedule> schedules = new ArrayList<>();

    private static final String EDUCATOR_XLSX_FILE_PATH = "./files/excel_files/educator.xlsx";

    public Educator(String id, String name, ArrayList<Schedule> schedules) {
        this.id = id;
        this.name = name;
        this.schedules = schedules;
    }

    public static ArrayList<Educator> readEducators() throws IOException, InvalidFormatException {
        ArrayList<Educator> educators = new ArrayList<>();

        Workbook workbook = WorkbookFactory.create(new File(EDUCATOR_XLSX_FILE_PATH));

        // Получение листа по индексу (первый лист)
        Sheet sheet = workbook.getSheetAt(0);

        // Создание DataFormatter для форматирования и получения значения ячейки как String
        DataFormatter dataFormatter = new DataFormatter();

        // цикл for-each для итерации по строкам и столбцам
        for (Row row: sheet) {
            if (row.getRowNum() != 0) {
                String id = dataFormatter.formatCellValue(row.getCell(0));
                String name = dataFormatter.formatCellValue(row.getCell(1));

                ArrayList<Schedule> schedules = Schedule.readSchedulesOfEducator(id);

                Educator educator = new Educator(id, name, schedules);
                educators.add(educator);
            }
        }

        // Закрытие workbook
        workbook.close();

        return educators;
    }

    public static void encodeEducators(ArrayList<Educator> educators) throws IOException, InvalidFormatException {
        JSONArray educatorsJSON = new JSONArray();
        for (Educator educator : educators) {
            JSONObject educatorJSON = new JSONObject();
            educatorJSON.put("id", educator.getId());
            educatorJSON.put("name", educator.getName());

            JSONArray schedulesJSON = new JSONArray();
            for (Schedule schedule : educator.getSchedules()) {
                JSONObject scheduleJSON = new JSONObject();
                scheduleJSON.put("id", schedule.getId());
                scheduleJSON.put("subject", Subject.getNameById(schedule.getSubjectId()));
                scheduleJSON.put("day", schedule.getDay());
                scheduleJSON.put("start_time", schedule.getStartTime());
                scheduleJSON.put("group", schedule.getGroup());

                schedulesJSON.put(scheduleJSON);
            }

            educatorJSON.put("schedule", schedulesJSON);

            educatorsJSON.put(educatorJSON);
        }

        String currentDir = System.getProperty("user.dir") + "/files/json_files";
        String filePath = Paths.get(currentDir, "educators.json").toString();

        try (FileWriter file = new FileWriter(filePath)) {
            file.write(educatorsJSON.toString(4)); // 4 - отступ для красивого форматирования
            System.out.println("JSON успешно сохранен в: " + filePath);
        } catch (IOException e) {
            System.out.println("Ошибка при сохранении файла: " + e.getMessage());
        }
    }

    public String getId() {
        return id;
    }

    public String getName() {
        return name;
    }

    public ArrayList<Schedule> getSchedules() {
        return schedules;
    }
}
