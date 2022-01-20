import com.github.junrar.Junrar;
import com.github.junrar.exception.RarException;
import net.lingala.zip4j.ZipFile;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.HashSet;
import java.util.Objects;
import java.util.Set;

public class MainClass {

    private static final String[] SPECIAL_SPLIT_CHARACTERS = {"/", ";", ",", "-", "\n", "hoặc", "|"};
    private static final String[] PHONE_COLUMN_CHARACTERS = {"thoai", "phone", "thoại", "đt", "sđt", "sdt", "động", "di dong", "tel", "mobile"};
    private static final String PHONE_NUMBER_REGEX = "^(0|\\+84)(\\s|\\.)?((3[2-9])|(5[689])|(7[06-9])|(8[1-689])|(9[0-46-9]))(\\d)(\\s|\\.)?(\\d{3})(\\s|\\.)?(\\d{3})$";
    private static final String ROOT_PATH = "/Users/macbook/Documents/Ads/";

    public static void main(String[] args) {
//        exportPhoneNumberToFile();
        exportPhoneNumberFromZipToFile();
    }

    public static void exportPhoneNumberFromZipToFile() {
        File zipDir = new File(ROOT_PATH + "file-zip");
        File unzipDir = new File(ROOT_PATH + "import-unzip");
        if (Objects.requireNonNull(zipDir.listFiles()).length > 0) {
            unzipDir(zipDir, unzipDir);
        }
    }

    public static void unzipDir(File zipDir, File unzipDir) {
        for (File file : Objects.requireNonNull(zipDir.listFiles())) {
            if (FilenameUtils.getExtension(file.getName()).equalsIgnoreCase("rar")) {
                try {
                    Junrar.extract(file, unzipDir);

                } catch (RarException | IOException e) {
                    System.out.println("rar error: " + file.getName() + ", error mess: " + e.getMessage());
                }
            } else if (FilenameUtils.getExtension(file.getName()).equalsIgnoreCase("zip")) {
                try {
                    ZipFile zipFile = new ZipFile(file);
                    if (zipFile.isEncrypted()) {
                        System.out.println("Zip file" + file.getName() + "has password");
                        return;
                    }
                    zipFile.extractAll(ROOT_PATH + "import-unzip");

                } catch (Exception e) {
                    System.out.println("zip error: " + file.getName() + ", error mess: " + e.getMessage());
                }
            }
            try {
                FileUtils.forceDelete(file);
            } catch (IOException e) {
                System.out.println("Can't delete file: " + file.getName() + ", error mess: " + e.getMessage());
            }
        }
    }


    private static void exportPhoneNumberToFile() {
        String importDirPath = ROOT_PATH + "import";
        File importDir = new File(importDirPath);
        File errorDir = new File(ROOT_PATH + "error");
        File noPhoneColumnDir = new File(ROOT_PATH + "no-phone-column");

        if (!importDir.exists() && importDir.mkdirs()) {
            System.out.println("Import folder is not created, we have created new one for you <3");
            return;
        }

        Set<File> lstImportFile = null;
        try {
            lstImportFile = listFilesUsingFileWalkAndVisitor(importDirPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (CollectionUtils.isEmpty(lstImportFile)) {
            System.out.println("Import folder is empty, you need to put an excel file in here: " + importDirPath);
            return;
        }

        for (File fExcel : lstImportFile) {
            if (fExcel.isFile() && fExcel.exists()) {
                System.out.println("Start reading file " + fExcel.getPath());
                Workbook excelWorkBook = null;
                FileInputStream fStreamExcel;
                try {
                    fStreamExcel = new FileInputStream(fExcel);
                } catch (FileNotFoundException e) {
                    continue;
                }

                try {
                    String extension = FilenameUtils.getExtension(fExcel.getName());
                    if (extension.equalsIgnoreCase("xls")) {
                        excelWorkBook = new HSSFWorkbook(fStreamExcel);
                    } else if (extension.equalsIgnoreCase("xlsx")) {
                        excelWorkBook = new XSSFWorkbook(fStreamExcel);
                    }
                } catch (Exception e) {
                    System.out.println("Some error when read file " + fExcel.getName() + " (error: " + e.getMessage() + "), coping to error folder\n");
                    copyFile(fExcel, errorDir);
                    continue;
                }

                if (Objects.nonNull(excelWorkBook)) {
                    File fileWriter = createOrReuseFile(ROOT_PATH);
                    Set<String> lstPhones = new HashSet<>();

                    for (Sheet sheet : excelWorkBook) {
                        System.out.println("Reading sheet " + sheet.getSheetName() + " ----");
                        int phoneColumnIndex = -1;
                        int flag = 0;
                        try {
                            // Run first 10 row to find phone column index
                            for (Row row : sheet) {
                                flag++;
                                for (Cell cell : row) {
                                    DataFormatter formatter = new DataFormatter();
                                    String phoneHeader = formatter.formatCellValue(cell).toLowerCase();
                                    if (phoneHeader.length() < 20 && StringUtils.isNotBlank(getWordsContains(phoneHeader, PHONE_COLUMN_CHARACTERS))) {
                                        phoneColumnIndex = cell.getColumnIndex();
                                        flag = 11;
                                        break;
                                    }
                                }
                                if (flag > 10) {
                                    break;
                                }
                            }
                        } catch (Exception e) {
                            System.out.println("Some error when get phoneColumnIndex: " + e.getMessage() + ", coping to error folder");
                            copyFile(fExcel, errorDir);
                            continue;
                        }

                        if (phoneColumnIndex == -1) {
                            System.out.println("Can't find phone's column or because sheet is empty, coping to no-phone-column folder");
                            copyFile(fExcel, noPhoneColumnDir);
                            continue;
                        }

                        for (Row row : sheet) {
                            if (isAllCellNull(row) && isAllCellNull(sheet.getRow(row.getRowNum() + 1))
                                    && isAllCellNull(sheet.getRow(row.getRowNum() + 2))) {
                                System.out.println("Sheet(" + sheet.getSheetName() + ") has stop at row " + row.getRowNum());
                                break;
                            }
                            try {
                                Cell cellPhone = row.getCell(phoneColumnIndex);
                                if (Objects.nonNull(cellPhone) && cellPhone.getCellType() != CellType.BLANK) {
                                    DataFormatter formatter = new DataFormatter();
                                    String phone = formatter.formatCellValue(cellPhone);
                                    if (StringUtils.isBlank(phone)) continue;

                                    String phoneNoWhiteSpace = StringUtils.deleteWhitespace(phone)
                                            .replace("+84", "0").replace(".", "");

                                    String wordsContains = getWordsContains(phoneNoWhiteSpace, SPECIAL_SPLIT_CHARACTERS);
                                    if (StringUtils.isNotBlank(wordsContains)) {
                                        for (String phoneGap : phoneNoWhiteSpace.split(wordsContains)) {
                                            addPhoneNumberToList(lstPhones, phoneGap);
                                        }
                                    } else {
                                        addPhoneNumberToList(lstPhones, phoneNoWhiteSpace);
                                    }
                                }
                            } catch (NullPointerException e) {
                                e.printStackTrace();
                            }
                        }
                    }

                    try {
                        if (CollectionUtils.isNotEmpty(lstPhones)) {
                            FileUtils.writeLines(fileWriter, lstPhones, true);
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }

                System.out.println("Finish reading file " + fExcel.getName() + "\n");
            }
        }
    }

    private static void addPhoneNumberToList(Set<String> lstPhones, String phone) {
        if (!phone.startsWith("0")) phone = "0" + phone;
        if (phone.matches(PHONE_NUMBER_REGEX)) {
            lstPhones.add(phone);
        }
    }

    private static File createOrReuseFile(String rootPath) {
        File dirExportFile = new File(rootPath + "/exported/");
        File fileWriter;
        File[] lstExportFile = dirExportFile.listFiles((dir, name) -> !name.equals(".DS_Store"));
        if (ArrayUtils.isEmpty(lstExportFile)) {
            fileWriter = new File(dirExportFile.getPath() + "/1.txt");
            try {
                if (fileWriter.createNewFile()) {
                    System.out.println("1.txt was created");
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            File lastFileModified = getLastModified(dirExportFile.getPath());
            if (countLineFast(lastFileModified.getPath()) < 900000) {
                System.out.println("Keep using file " + lastFileModified.getName() + " to write");
                return lastFileModified;
            }
            String fileName = FilenameUtils.removeExtension(lastFileModified.getName());

            String newTxtFile = (Integer.parseInt(fileName) + 1) + ".txt";
            fileWriter = new File(dirExportFile.getPath() + "/" + newTxtFile);
            try {
                if (fileWriter.createNewFile()) {
                    System.out.println(newTxtFile + " was created");
                }
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
        return fileWriter;
    }

    public static File getLastModified(String directoryFilePath) {
        File directory = new File(directoryFilePath);
        File[] files = directory.listFiles(File::isFile);
        long lastModifiedTime = Long.MIN_VALUE;
        File chosenFile = null;

        if (files != null) {
            for (File file : files) {
                if (file.lastModified() > lastModifiedTime) {
                    chosenFile = file;
                    lastModifiedTime = file.lastModified();
                }
            }
        }

        return chosenFile;
    }

    public static long countLineFast(String fileName) {

        long lines = 0;

        try (InputStream is = new BufferedInputStream(new FileInputStream(fileName))) {
            byte[] c = new byte[1024];
            int count = 0;
            int readChars = 0;
            boolean endsWithoutNewLine = false;
            while ((readChars = is.read(c)) != -1) {
                for (int i = 0; i < readChars; ++i) {
                    if (c[i] == '\n')
                        ++count;
                }
                endsWithoutNewLine = (c[readChars - 1] != '\n');
            }
            if (endsWithoutNewLine) {
                ++count;
            }
            lines = count;
        } catch (IOException e) {
            e.printStackTrace();
        }

        return lines;
    }

    public static String getWordsContains(String inputString, String[] items) {
        for (String item : items) {
            if (inputString.contains(item)) {
                return item;
            }
        }
        return null;
    }

    public static Set<File> listFilesUsingFileWalkAndVisitor(String dir) throws IOException {
        Set<File> fileList = new HashSet<>();
        Files.walkFileTree(Paths.get(dir), new SimpleFileVisitor<>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                String fileName = file.getFileName().toString();
                String extension = FilenameUtils.getExtension(fileName).toLowerCase();
                if (extension.contains("zip") || extension.contains("rar")) {
                    copyFile(file.toFile(), new File(ROOT_PATH + "file-zip"));
                }
                if (!Files.isDirectory(file) && FilenameUtils.getExtension(fileName).toLowerCase().contains("xls")
                        && !fileName.equals(".DS_Store") && !fileName.startsWith("~$")) {
                    fileList.add(file.toFile());
                }
                return FileVisitResult.CONTINUE;
            }
        });
        return fileList;
    }

    public static boolean isAllCellNull(Row row) {
        if (Objects.isNull(row)) {
            return true;
        }
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK)
                return false;
        }
        return true;
    }

    private static void copyFile(File file, File dirFile) {
        try {
            FileUtils.copyFileToDirectory(file, dirFile);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

}
