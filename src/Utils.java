import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public final class Utils {

    public static XSSFWorkbook getWorkbookSafe(final File file, File outTempFile) throws IOException, InvalidFormatException {
        outTempFile = File.createTempFile("ExcelExchangeUSD_ARS", getFileNameExt(file.getName()));
        outTempFile.deleteOnExit();

        try(
            InputStream in = new BufferedInputStream(
                    new FileInputStream(file));
            OutputStream out = new BufferedOutputStream(
                    new FileOutputStream(outTempFile))) {

            byte[] buffer = new byte[1024];
            int lengthRead;
            while ((lengthRead = in.read(buffer)) > 0) {
                out.write(buffer, 0, lengthRead);
                out.flush();
            }
        }

        return new XSSFWorkbook(outTempFile);
    }

    public static String getFileNameExt(String fileName){
        String[] split = fileName.split("\\.");
        return split[split.length-1];
    }

    public static String addStringToEndFileName(String fileName, String str){
        String ext = getFileNameExt(fileName);
        return fileName.replace("."+ext, str+"."+ext);
    }
}
