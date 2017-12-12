package com.example.sadi.excelreadwrite;

import android.app.Activity;
import android.content.ActivityNotFoundException;
import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.Iterator;
import java.util.Locale;

import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelActivity extends Activity {
    /** Called when the activity is first created. */
    Intent intent ;
    Button btnExplore;
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_excel);

//        View writeButton = findViewById(R.id.write);
//        View readButton = findViewById(R.id.read);
//        View writeExcelButton = findViewById(R.id.writeExcel);
//        View readExcelButton = findViewById(R.id.readExcel);
        btnExplore = (Button) findViewById(R.id.btnExplore);
        btnExplore.setOnClickListener(new OnClickListener() {
            @Override
            public void onClick(View view) {
                intent = new Intent(Intent.ACTION_GET_CONTENT);
                intent.setType("*/*");
                startActivityForResult(intent, 7);
            }
        });

        //readFile(getApplicationContext(),"Book.xlsx");
        //txtFileRead();
       // readViewExcel();
       // createReadXlsx();

    }
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        // TODO Auto-generated method stub

        switch(requestCode){

            case 7:

                if(resultCode==RESULT_OK){

                    String PathHolder = data.getData().getPath();

                    Toast.makeText(ExcelActivity.this, PathHolder , Toast.LENGTH_LONG).show();

                }
                break;

        }
    }
//    private void createReadXlsx() {
//
//        String Fnamexls="testfile"  + ".xls";
//        File sdCard = Environment.getExternalStorageDirectory();
//        File directory = new File (sdCard.getAbsolutePath() + "/xlfolder");
//        directory.mkdirs();
//        File file = new File(directory, Fnamexls);
//
//        WorkbookSettings wbSettings = new WorkbookSettings();
//
//        wbSettings.setLocale(new Locale("en", "EN"));
//
//        WritableWorkbook workbook;
//        try {
//            int a = 1;
//            workbook = Workbook.createWorkbook(file, wbSettings);
//            //workbook.createSheet("Report", 0);
//            WritableSheet sheet = workbook.createSheet("First Sheet", 0);
//            Label label = new Label(0, 2, "SECOND");
//            Label label1 = new Label(0,1,"first");
//            Label label0 = new Label(0,0,"HEADING");
//            Label label3 = new Label(1,0,"Heading2");
//            Label label4 = new Label(1,1,String.valueOf(a));
//            try {
//                sheet.addCell(label);
//                sheet.addCell(label1);
//                sheet.addCell(label0);
//                sheet.addCell(label4);
//                sheet.addCell(label3);
//            } catch (RowsExceededException e) {
//                // TODO Auto-generated catch block
//                e.printStackTrace();
//            } catch (WriteException e) {
//                // TODO Auto-generated catch block
//                e.printStackTrace();
//            }
//
//
//            workbook.write();
//            try {
//                workbook.close();
//            } catch (WriteException e) {
//                // TODO Auto-generated catch block
//                e.printStackTrace();
//            }
//            //createExcel(excelSheet);
//        } catch (IOException e) {
//            // TODO Auto-generated catch block
//            e.printStackTrace();
//        }
//    }

    private void readViewExcel() {
//        File file = new File(Environment.getExternalStorageDirectory()
//                + "Book.xlsx");
//        if (file .exists())
//        {
//            Uri path = Uri.fromFile(file );
//            Intent pdfIntent = new Intent(Intent.ACTION_VIEW);
//            pdfIntent.setDataAndType(path , "application/vnd.ms-excel");
//            pdfIntent.setFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP);
//            try
//            {
//                startActivity(pdfIntent ); }
//            catch (ActivityNotFoundException e)
//            {
//                Toast.makeText(ExcelActivity.this,"Please install MS-Excel app to view the file.",
//                        Toast.LENGTH_SHORT).show();
//            }
//        }
    }


    private void txtFileRead() {

        File sdcard = Environment.getExternalStorageDirectory();

//Get the text file
        File file = new File(sdcard,"dimens.txt");

//Read text from file
        StringBuilder text = new StringBuilder();

        try {
            BufferedReader br = new BufferedReader(new FileReader(file));
            String line;

            while ((line = br.readLine()) != null) {
                text.append(line);
                text.append('\n');
            }

            Toast.makeText(this, ""+text, Toast.LENGTH_SHORT).show();
            br.close();
        }
        catch (IOException e) {
            //You'll need to add proper error handling here
        }
    }

    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.write:
                //saveFile(this,"myFile.txt");
                break;
            case R.id.read:
                //readFile(this,"myFile.txt");
                break;
            case R.id.writeExcel:
               // saveExcelFile(this,"myExcel.xls");
                break;
            case R.id.readExcel:
               // readExcelFile(this,"Book.xlsx");
                break;
        }
    }

    private static boolean saveFile(Context context, String fileName) {

        // check if available and not read only
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.w("FileUtils", "Storage not available or read only");
            return false;
        }

        // Create a path where we will place our List of objects on external storage
        File file = new File(context.getExternalFilesDir(null), fileName);
        PrintStream p = null; // declare a print stream object
        boolean success = false;

        try {
            OutputStream os = new FileOutputStream(file);
            // Connect print stream to the output stream
            p = new PrintStream(os);
            p.println("This is a TEST");
            Log.w("FileUtils", "Writing file" + file);
            success = true;
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != p)
                    p.close();
            } catch (Exception ex) {
            }
        }

        return success;
    }

    private static void readFile(Context context, String filename) {

        if (!isExternalStorageAvailable() || isExternalStorageReadOnly())
        {
            Log.w("FileUtils", "Storage not available or read only");
            return;
        }

        FileInputStream fis = null;

        try
        {
            String file = context.getFilesDir() + "/" + filename;
            fis = new FileInputStream(file);
            // Get the object of DataInputStream
            DataInputStream in = new DataInputStream(fis);
            BufferedReader br = new BufferedReader(new InputStreamReader(in));
            String strLine;
            //Read File Line By Line
            while ((strLine = br.readLine()) != null) {
                Log.e("FileUtils", "File data: " + strLine);
                Toast.makeText(context, "File Data: " + strLine , Toast.LENGTH_SHORT).show();
            }
            in.close();
        }
        catch (Exception ex) {
            Log.e("FileUtils", "failed to load file", ex);
        }
        finally {
            try {if (null != fis) fis.close();} catch (IOException ex) {}
        }

        return;
    }

    private static boolean saveExcelFile(Context context, String fileName) {

        // check if available and not read only
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.w("FileUtils", "Storage not available or read only");
            return false;
        }

        boolean success = false;

        //New Workbook
        Workbook wb = new HSSFWorkbook();

        Cell c = null;

        //Cell style for header row
        CellStyle cs = wb.createCellStyle();
        cs.setFillForegroundColor(HSSFColor.LIME.index);
        cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("myOrder");

        // Generate column headings
        Row row = sheet1.createRow(0);

        c = row.createCell(0);
        c.setCellValue("Item Number");
        c.setCellStyle(cs);

        c = row.createCell(1);
        c.setCellValue("Quantity");
        c.setCellStyle(cs);

        c = row.createCell(2);
        c.setCellValue("Price");
        c.setCellStyle(cs);

        sheet1.setColumnWidth(0, (15 * 500));
        sheet1.setColumnWidth(1, (15 * 500));
        sheet1.setColumnWidth(2, (15 * 500));

        // Create a path where we will place our List of objects on external storage
        File file = new File(context.getExternalFilesDir(null), fileName);
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
            success = true;
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }

        return success;
    }

    private static void readExcelFile(Context context, String filename) {

        if (!isExternalStorageAvailable() || isExternalStorageReadOnly())
        {
            Log.w("FileUtils", "Storage not available or read only");
            return;
        }

        try{
            // Creating Input Stream
            File file = new File(context.getExternalFilesDir(null), filename);
            FileInputStream myInput = new FileInputStream(file);

            // Create a POIFSFileSystem object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            // Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

            // Get the first sheet from workbook
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);

            /** We now need something to iterate through the cells.**/
            Iterator<Row> rowIter = mySheet.rowIterator();

            while(rowIter.hasNext()){
                HSSFRow myRow = (HSSFRow) rowIter.next();
                Iterator<Cell> cellIter = myRow.cellIterator();
                while(cellIter.hasNext()){
                    HSSFCell myCell = (HSSFCell) cellIter.next();
                    Log.w("FileUtils", "Cell Value: " +  myCell.toString());
                    Toast.makeText(context, "cell Value: " + myCell.toString(), Toast.LENGTH_SHORT).show();
                }
            }
        }catch (Exception e){e.printStackTrace(); }

        return;
    }

    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) {
            return true;
        }
        return false;
    }


}
