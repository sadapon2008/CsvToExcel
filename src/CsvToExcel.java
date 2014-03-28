import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.io.FileOutputStream;

import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

import au.com.bytecode.opencsv.CSVReader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class CsvToExcel {

	public static void main(String[] args) {
		new CsvToExcel().start(args);

	}

	public void start(String[] args) {
        String filename_datacsv = "";  // データCSVファイル
        String filename_paramcsv = ""; // パラメータCSVファイル
        String filename_templatexlsx = "";      // 入力テンプレートxlsxファイル
        String filename_outxlsx = "";      // 出力xlsxファイル
			
        // コマンドライン引数を解析する
        Options opts = new Options();
        opts.addOption("d", "datacsv", true, "input data csv file");
        opts.addOption("p", "paramcsv", true, "input parameter csv file");
        opts.addOption("t", "templatexlsx", true, "input template xlsx file");
        opts.addOption("o", "outxlsx", true, "output xlsx file");
        BasicParser parser = new BasicParser();
        CommandLine cl;
        HelpFormatter help = new HelpFormatter();
        try {
        	// 解析する
        	cl = parser.parse(opts, args);
        
        	filename_datacsv = cl.getOptionValue("d");
        	if(filename_datacsv == null) {
        		throw new ParseException("");
        	}
        	filename_paramcsv = cl.getOptionValue("p");
        	if(filename_paramcsv == null) {
        		throw new ParseException("");                                
        	}
        	filename_templatexlsx = cl.getOptionValue("t");
        	if(filename_templatexlsx == null) {
        		throw new ParseException("");                                
        	}
        	filename_outxlsx = cl.getOptionValue("o");
        	if(filename_outxlsx == null) {
        		throw new ParseException("");                
        	}
        }catch (ParseException e) {
        	help.printHelp("CsvToExcel", opts);
        	System.exit(1);
        }
    
        // パラメータCSVファイルからパラメータを読み込み
        String param_sheet_name = "";
        int param_template_sheet_index = -1;
        int param_header_row_count = -1;
        int param_footer_row_count = -1;
    	int[] headerType = null;
    	int[] bodyType = null;
    	int[] footerType = null;
        
        try {
        	FileInputStream input=new FileInputStream(filename_paramcsv);
        	InputStreamReader inReader=new InputStreamReader(input, "UTF-8");
        	CSVReader reader = new CSVReader(inReader,',','"');
        	String [] nextLine;
        	
        	// 1行目は設定行
        	if (((nextLine = reader.readNext()) == null) 
        		|| (nextLine.length < 4)) {
            	System.out.println("error in parsing paramcsv");
            	System.exit(1);
        	}
        	param_sheet_name = nextLine[0]; // 出力シート名
        	param_template_sheet_index = Integer.parseInt(nextLine[1]); // 入力テンプレートのシートインデックス
        	param_header_row_count = Integer.parseInt(nextLine[2]); // ヘッダー行数
        	param_footer_row_count = Integer.parseInt(nextLine[3]); // フッター行数
        	
        	// 2行目はヘッダーのデータタイプ行
        	if (((nextLine = reader.readNext()) == null)
        		|| (nextLine.length < 1)) {
            	System.out.println("error in parsing paramcsv");
            	System.exit(1);        		
        	}
        	headerType = new int[nextLine.length];
        	for(int i = 0; i < nextLine.length; i++) {
        		if(nextLine[i].equals("formula")) {
        			headerType[i] = Cell.CELL_TYPE_FORMULA;
        		} else if(nextLine[i].equals("numeric")) {
        			headerType[i] = Cell.CELL_TYPE_NUMERIC;
        		} else {
        			headerType[i] = Cell.CELL_TYPE_STRING;            			
        		}
        	}
        	
        	// 3行目はボディのデータタイプ行
        	if (((nextLine = reader.readNext()) == null)
        		|| (nextLine.length < 1)) {
            	System.out.println("error in parsing paramcsv");
            	System.exit(1);        		
        	}
        	bodyType = new int[nextLine.length];
        	for(int i = 0; i < nextLine.length; i++) {
        		if(nextLine[i].equals("formula")) {
        			bodyType[i] = Cell.CELL_TYPE_FORMULA;
        		} else if(nextLine[i].equals("numeric")) {
        			bodyType[i] = Cell.CELL_TYPE_NUMERIC;
        		} else {
        			bodyType[i] = Cell.CELL_TYPE_STRING;            			
        		}
        	}

        	// 4行目はフッターのデータタイプ行
        	if (((nextLine = reader.readNext()) == null)
        		|| (nextLine.length < 1)) {
            	System.out.println("error in parsing paramcsv");
            	System.exit(1);        		
        	}
        	footerType = new int[nextLine.length];
        	for(int i = 0; i < nextLine.length; i++) {
        		if(nextLine[i].equals("formula")) {
        			footerType[i] = Cell.CELL_TYPE_FORMULA;
        		} else if(nextLine[i].equals("numeric")) {
        			footerType[i] = Cell.CELL_TYPE_NUMERIC;
        		} else {
        			footerType[i] = Cell.CELL_TYPE_STRING;            			
        		}
        	}
        	reader.close();            
        } catch(Exception e) {
        	System.out.println("error in parsing paramcsv");
        	System.out.println(e.toString());
        	System.exit(1);
        }

        // データCSVの行数をカウントする
        int num_row = 0;
        try {
        	FileInputStream input=new FileInputStream(filename_datacsv);
        	InputStreamReader inReader=new InputStreamReader(input, "UTF-8");
        	CSVReader reader = new CSVReader(inReader,',','"');
        	while (reader.readNext() != null) {
        		num_row++;
        	}
        	reader.close();            
        } catch(Exception e) {
        	System.out.println("error in parsing datacsv");
        	System.out.println(e.toString());
        	System.exit(1);
        }
       	
        int num_col = 0;
        try {
        	// 入力テンプレートファイルを開く
            OPCPackage pkg = OPCPackage.open(new File(filename_templatexlsx));
            XSSFWorkbook wb_in = new XSSFWorkbook(pkg);
        	
            int num_sheet_in = wb_in.getNumberOfSheets();
            
            // テンプレートシート
        	XSSFSheet sheet_in = wb_in.getSheetAt(param_template_sheet_index);
        	
        	// 出力するシートを新規作成
        	XSSFSheet sheet_out = wb_in.createSheet(param_sheet_name);
        	
        	// データCSVを読み込んでデータを作成していく
          	FileInputStream input=new FileInputStream(filename_datacsv);
           	InputStreamReader inReader=new InputStreamReader(input, "UTF-8");
           	CSVReader reader = new CSVReader(inReader,',','"');
            	
            int n = 0;
            	
           	String [] nextLine;
           	while ((nextLine = reader.readNext()) != null) {
           		// 行を新規作成
           		XSSFRow row_out = sheet_out.createRow(n);
           		n++;
           		int row_in_index;
           		boolean is_header = false;
           		boolean is_body = false;
           		boolean is_footer = false;
           		if(n <= param_header_row_count) {
           			// ヘッダー行
           			is_header = true;
           			row_in_index = n-1;
           		} else if((param_header_row_count < n) && (n <= num_row - param_footer_row_count)) {
           			// データ行
           			is_body = true;
           			row_in_index = param_header_row_count;
           		} else {
           			// フッター行
           			is_footer = true;
           			row_in_index = param_header_row_count+(param_footer_row_count-(num_row-n));
           		}
           		// テンプレートシートの行
           		XSSFRow row_in = sheet_in.getRow(row_in_index);
           		
           		// 列数をカウントしておく
           		if(num_col < row_in.getLastCellNum()+1) {
           			num_col = row_in.getLastCellNum()+1;
           		}
           		
           		// 行の高さを揃える
       			row_out.setHeightInPoints(row_in.getHeightInPoints());
       			for(int c = 0; c < nextLine.length; c++) {
       				// セルを新規作成
       				XSSFCell cell_out = row_out.createCell(c);
       				XSSFCell cell_in = row_in.getCell(c);
       				// 書式のコピー
       				if(cell_in != null) {
       					cell_out.setCellStyle(cell_in.getCellStyle());
       				}
       				// パラメータCSVで設定した形式に基づいてデータをセットする
       				int cell_type = cell_out.getCellType();
       				if(is_header && (c < headerType.length)) {
       					cell_type = headerType[c];
       				} else if(is_body && (c < bodyType.length)) {
       					cell_type = bodyType[c];
       				} else if(is_footer && (c < footerType.length)) {
       					cell_type = footerType[c];
       				}
       				if(cell_type == Cell.CELL_TYPE_FORMULA) {
       					cell_out.setCellFormula(nextLine[c]);
       				} else if(cell_type == Cell.CELL_TYPE_NUMERIC) {
       					cell_out.setCellValue(Double.parseDouble(nextLine[c]));
       				} else {
       					cell_out.setCellValue(nextLine[c]);
       				}
            	}
           	}
            reader.close();            

            // 列幅を揃える
            for(int c = 0; c < num_col; c++) {
            	sheet_out.setColumnWidth(c, sheet_in.getColumnWidth(c));
            }
            
            // セルの結合を揃える
            for(int i = 0; i < sheet_in.getNumMergedRegions(); i++) {
            	CellRangeAddress ra = sheet_in.getMergedRegion(i);
            	if((ra.getFirstRow() < param_header_row_count)
            		&& (ra.getLastRow() < param_header_row_count)) {
            		// ヘッダー
            		sheet_out.addMergedRegion(new CellRangeAddress(ra.getFirstRow(), ra.getLastRow(), ra.getFirstColumn(), ra.getLastColumn()));
            	} else if((ra.getFirstRow() == param_header_row_count)
            			&& (ra.getLastRow() == param_header_row_count)) {
            		// ボディ
            		for(int r = param_header_row_count; r < (num_row-param_footer_row_count); r++) {
                		sheet_out.addMergedRegion(new CellRangeAddress(r, r, ra.getFirstColumn(), ra.getLastColumn()));            			
            		}
            	} else if((ra.getFirstRow() >= param_header_row_count+1)
            		&& (ra.getLastRow() >= param_header_row_count+1)) {
            		// フッター
            		int first_row = (num_row-param_footer_row_count) + (ra.getFirstRow()-(param_header_row_count+1));
            		int last_row = (num_row-param_footer_row_count) + (ra.getLastRow()-(param_header_row_count+1));
            		sheet_out.addMergedRegion(new CellRangeAddress(first_row, last_row, ra.getFirstColumn(), ra.getLastColumn()));            		
            	}
            }
            
            // 他のシートを削除する
            for(int i = num_sheet_in-1; i >= 0; i--) {
            	wb_in.removeSheetAt(param_template_sheet_index);
            }
            wb_in.setActiveSheet(0);
            
            // ファイルを別名で保存する
        	FileOutputStream fileOut = new FileOutputStream(filename_outxlsx);
        	wb_in.write(fileOut);
        	fileOut.close();
        } catch(Exception e) {
        	System.out.println(e.getMessage());
        }
	}
}