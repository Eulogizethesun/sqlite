import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main
{
    static List<String> type=new ArrayList<String>();
    static List<String> category=new ArrayList<String>();
    static int colnum=0;static int rownum=0;
    static String sheet_name=null;
    static Workbook wb = null;
    static Sheet sheet = null;
    static String name_db=null;static String name_excel=null;static String name_excel_sheet=null;
    static Connection c = null;
    static Statement stmt = null;
    public static void main( String args[] )
    {
        if (args.length==4) {
            name_db=args[0];name_excel=args[1];name_excel_sheet=args[2];sheet_name=args[3];
        }
        else if (args.length==3) {
            name_db=args[0];name_excel=args[1];name_excel_sheet=args[2];sheet_name=args[2];
        }
        else if (args.length==2) {
            name_db=args[0];name_excel=args[1];name_excel_sheet=null;sheet_name=null;
        }
        String filePath = name_excel;
        String sql="";
        int i=0,j=0;
        try {
            Class.forName("org.sqlite.JDBC");
            c = DriverManager.getConnection("jdbc:sqlite:"+args[0]);
            System.out.println("Opened database successfully");
            stmt = c.createStatement();

            wb = readExcel(filePath);
            if (wb != null) {
                //获取第一个sheet
                if (name_excel_sheet == null) {
                    sheet = wb.getSheetAt(0);
                    name_excel_sheet = sheet.getSheetName();
                } else {
                    sheet = wb.getSheet(name_excel_sheet);
                }
                if (sheet_name == null) {
                    sheet_name = sheet.getSheetName();
                }

                DatabaseMetaData d = c.getMetaData();
                ResultSet r = d.getTables(null, null, sheet_name, null);
                if (r.next()) {
                    r.close();
                    sql = "DROP TABLE " + sheet_name;
                    stmt.executeUpdate(sql);
                }
            }


            preread();

            sql = "CREATE TABLE "+sheet_name+" " +
                    "(Line INTEGER PRIMARY KEY AUTOINCREMENT)" ;
            stmt.executeUpdate(sql);


            for (j=0;j<colnum;j++)
            {
                sql="alter table "+sheet_name+" add column "+category.get(j)+" "+type.get(j);
                stmt.execute(sql);
            }

            for (i=1;i<rownum;i++)
            {
                sql="INSERT INTO "+sheet_name+" ";
                sql=sql+"VALUES (NULL,";
                for (j=0;j<colnum;j++) {
                    if (type.get(j).charAt(0)=='c'){
                    sql = sql + "'"+sheet.getRow(i).getCell(j).getStringCellValue()+"'";
                    }
                    else if (type.get(j).charAt(0)=='i'){
                        sql = sql + String.valueOf((int)(sheet.getRow(i).getCell(j).getNumericCellValue()));
                    }
                    else if (type.get(j).charAt(0)=='r') {
                        sql = sql + String.valueOf(sheet.getRow(i).getCell(j).getNumericCellValue());
                    }
                    if (j==colnum-1) {
                        sql = sql + ")";
                    }
                    else {
                        sql = sql + ",";
                    }
                }
                sql+=";";

                stmt.execute(sql);
            }
            System.out.println("表结构");
            DatabaseMetaData m_DBMetaData = c.getMetaData();
            String columnName;
            String columnType;
            ResultSet colRet = m_DBMetaData.getColumns(null,"%", sheet_name,"%");
            while(colRet.next()) {
                columnName = colRet.getString("COLUMN_NAME");
                columnType = colRet.getString("TYPE_NAME");
                int datasize = colRet.getInt("COLUMN_SIZE");
                int digits = colRet.getInt("DECIMAL_DIGITS");
                int nullable = colRet.getInt("NULLABLE");
                System.out.println(columnName+" "+columnType+" COLUMN_SIZE:"+datasize+" DECIMAL_DIGITS:"+digits+" NULLABLE:"+ nullable);
            }


            sql = "select * from "+sheet_name;
            System.out.println("行数：");
            int rowCount=0;
            ResultSet rs = stmt.executeQuery(sql);
            while(rs.next()){
                rowCount = rs.getInt(1);
            }

            System.out.println(rowCount);



            stmt.close();
            c.close();
        } catch ( Exception e ) {
            System.err.println( e.getClass().getName() + ": " + e.getMessage() );
            System.exit(0);
        }

    }

    public static void preread() {
        Row row = null;
        String sql=null;
        if (wb != null) {
            //获取最大行数
            rownum = sheet.getPhysicalNumberOfRows();
            row = sheet.getRow(0);
            colnum = row.getPhysicalNumberOfCells();
            row=sheet.getRow(1);
            Cell cell=row.getCell(0);
            int i=0,j=0;
            for (j=0;j<colnum;j++)//判断字段类型
            {
                if (sheet.getRow(1).getCell(j).getCellType()==Cell.CELL_TYPE_NUMERIC)
                {
                    boolean point=false;
                    for (i = 1; i<rownum; i++)
                    {
                        cell = sheet.getRow(i).getCell(j);//i行j列
                        double value=cell.getNumericCellValue();
                        if((value-(int)value)!=0)
                        {
                            point=true;break;
                        }
                    }
                    if (point){
                        type.add("real");
                    }
                    else {
                        type.add("int");
                    }
                }
                else
                {
                    int l=0;
                    for (i = 1; i<rownum; i++)
                    {
                        cell = sheet.getRow(i).getCell(j);//i行j列
                        String value=cell.getStringCellValue();
                        l=(value.length()>l)?value.length():l;
                    }
                    type.add("char"+"("+String.valueOf(l)+")");
                }
            }
            for (j=0;j<colnum;j++)
            {
                category.add(j,sheet.getRow(0).getCell(j).getStringCellValue());
            }
        }
    }

    public static Workbook readExcel(String filePath){
        Workbook wb = null;
        if(filePath==null){
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                return wb = new HSSFWorkbook(is);
            }else if(".xlsx".equals(extString)){
                return wb = new XSSFWorkbook(is);
            }else{
                return wb = null;
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

}

