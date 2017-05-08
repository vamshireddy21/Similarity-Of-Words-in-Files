import jdk.nashorn.internal.objects.NativeArray;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Main {
    public static void main(String[] args) throws IOException {
        PDFManager pdfManager = new PDFManager();
        String pdfFilesDirectoryPath = "D:\\GRS Project 2\\Downloaded PDF Files";
        ArrayList<HashMap<String, Integer>> corpusNGramsList = new ArrayList<>();
        for (int index = 0; index <= 5; index++) {
            corpusNGramsList.add(index,new HashMap<String, Integer>());
        }

        final ArrayList<File> fileNamesInTheDirectory = getFileNamesInTheDirectory(pdfFilesDirectoryPath);
        for (File pdfFile : fileNamesInTheDirectory) {
            pdfManager.setFilePath(pdfFile.getAbsolutePath());
            String document = pdfManager.toText();
            if (document == "") {
                continue;
            }
            try (PrintStream out = new PrintStream(new FileOutputStream("D:\\GRS Project 2\\Downloaded PDF Files\\sample" + pdfFile.getName().replace("pdf", "txt")))) {
                out.print(document);
            }
            System.out.println(document);
            document = document.replaceAll("[!?,\"'\\u0022\\u005c()â€œâ€�.-:/]", "");
            String[] words = document.split("\\s+");
            ArrayList<HashMap<String, Integer>> nGramsMapList = new ArrayList<>();
            for (int n = 1; n <= 6; n++) {
                HashMap<String, Integer> nGramsMap = new HashMap<String, Integer>();

                for (int i = n-1; i < words.length; i++) {
                    StringBuilder nGram= new StringBuilder();
                    for (int j = n-1; j >= 0; j--) {
                        nGram = nGram.append(words[i-j].toLowerCase()).append("|");
                    }

                    System.out.println(nGram);
                    if (nGramsMap.containsKey(nGram.toString())) {
                        Integer count = nGramsMap.get(nGram.toString());
                        Integer countAll = corpusNGramsList.get(n-1).get(nGram.toString());
                        corpusNGramsList.get(n-1).put(nGram.toString(), countAll + 1);
                        nGramsMap.put(nGram.toString(), count + 1);
                    } else {
                        nGramsMap.put(nGram.toString(), 1);
                    }
                    if (corpusNGramsList.get(n-1).containsKey(nGram.toString())) {
                        Integer countAll = corpusNGramsList.get(n-1).get(nGram.toString());
                        corpusNGramsList.get(n-1).put(nGram.toString(), countAll + 1);
                    } else {
                        corpusNGramsList.get(n-1).put(nGram.toString(), 1);
                    }

                }
                nGramsMapList.add(n-1,nGramsMap);
            }
            createExcelSheetForFile(pdfFile.getName().replace("pdf", "xlsx"), nGramsMapList);
            serializeHashMap(nGramsMapList.get(5), pdfFile.getName().replace("pdf", "ser"));
        }
        createExcelSheetForFile("corpus.xlsx", corpusNGramsList);
    }

    private static void createExcelSheetForFile(String fileName, ArrayList<HashMap<String, Integer>> nGramsMapsList) {
        Workbook wb = new XSSFWorkbook();
        int n = 1;
        for (HashMap<String,Integer> nGramsMap : nGramsMapsList) {
            Sheet sheet = wb.createSheet(n + "-gram");
            n++;
            int rowCount = 1;
            for (String key : nGramsMap.keySet()) {
                Row row = sheet.createRow(rowCount);
                row.createCell(0).setCellValue(key);
                row.createCell(1).setCellValue(nGramsMap.get(key));
                rowCount++;
                System.out.println(key + " " + nGramsMap.get(key));
            }
        }
        try (PrintStream out = new PrintStream(new FileOutputStream("/Users/yash/GradProjFiles/GradProjExcel/" + fileName))) {
            wb.write(out);
        } catch (IOException e) {
                e.printStackTrace();
        }
    }

    private static void serializeHashMap(HashMap<String, Integer> hashMap, String fileName) {
        try (FileOutputStream fos = new FileOutputStream("/Users/yash/GradProjFiles/GradProjHashMaps/" + fileName) ; ObjectOutputStream oos = new ObjectOutputStream(fos)) {
            oos.writeObject(hashMap);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static ArrayList<File> getFileNamesInTheDirectory(String directoryAbsolutePath) {
        File folder = new File(directoryAbsolutePath);
        File[] listOfFilesAndDirectories = folder.listFiles();
        ArrayList<File> pdfFiles = new ArrayList<File>();
        for (int i = 0; i < listOfFilesAndDirectories.length; i++) {
            if (listOfFilesAndDirectories[i].isFile()) {
                System.out.println("File " + listOfFilesAndDirectories[i].getAbsolutePath());
                if (listOfFilesAndDirectories[i].getName().endsWith(".pdf")){
                    pdfFiles.add(listOfFilesAndDirectories[i]);
                }
            } else if (listOfFilesAndDirectories[i].isDirectory()) {
                System.out.println("Directory " + listOfFilesAndDirectories[i].getName());
            }
        }
        return pdfFiles;
    }
}

