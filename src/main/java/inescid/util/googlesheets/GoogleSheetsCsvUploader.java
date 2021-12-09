package inescid.util.googlesheets;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;

import org.apache.commons.collections4.keyvalue.DefaultMapEntry;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.FileUtils;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.Drive.Permissions;
import com.google.api.services.drive.model.Permission;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Get;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Clear;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Update;
import com.google.api.services.sheets.v4.model.AddSheetRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.ExtendedValue;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.SpreadsheetProperties;
import com.google.api.services.sheets.v4.model.UpdateSheetPropertiesRequest;
import com.google.api.services.sheets.v4.model.ValueRange;

public class GoogleSheetsCsvUploader {
	public static void update(String spreadsheetId, File csvFile) throws IOException {
		update(spreadsheetId, sheetTitleFromFileName(csvFile), csvFile);
	}
	
	public static void update(String spreadsheetId, String sheetTitle, File csvFile) throws IOException {
		Sheets service = GoogleApi.getSheetsService();

		Get get = service.spreadsheets().get(spreadsheetId);
		get.setFields("sheets.properties");
		
		boolean sheetExists=false;
		Spreadsheet result = get.execute();
		for(Sheet sheet : result.getSheets()) {
			SheetProperties sheetProps = sheet.getProperties();
			if(!sheetProps.getTitle().equals(sheetTitle))
				continue;
			sheetExists=true;
			String range=GoogleSheetsApi.makeRangeExpression(sheetTitle, sheetProps.getGridProperties().getRowCount(), sheetProps.getGridProperties().getColumnCount());
			ClearValuesRequest requestBody = new ClearValuesRequest();
			Clear clear = service.spreadsheets().values().clear(spreadsheetId, range, requestBody);
			clear.execute();
			break;
//			System.out.println("Cleared: "+ executeClear.getClearedRange()/*row with blanks*/+" range");
		}
		if(!sheetExists) {
			GoogleSheetsApi.addSheet(spreadsheetId, sheetTitle);
		}
		 
		Entry<ValueRange, Integer> createValues = createValues(csvFile);
		ValueRange vRange=createValues.getKey();
		int cols=createValues.getValue();
		int rows=vRange.getValues().size();
		
		Update append = service.spreadsheets().values().update(spreadsheetId, GoogleSheetsApi.makeRangeExpression(sheetTitle, rows, cols) , vRange);
		append.setValueInputOption("RAW");
		append.execute();
	}
	

	private static Entry<ValueRange, Integer> createValues(File csvFile) throws IOException {
		int cols=0;
        List<List<Object>> vals=new ArrayList<List<Object>>();
        CSVParser parser=CSVParser.parse(FileUtils.readFileToString(csvFile, "UTF-8"), CSVFormat.DEFAULT);
		for(Iterator<CSVRecord> it = parser.iterator() ; it.hasNext() ; ) {
			CSVRecord rec = it.next();
			cols=Math.max(cols, rec.size());
			
			List<Object> recVals=new ArrayList<>();
			for(String v:rec) {
				if(v.length()>=5000) {
					recVals.add(v.substring(0, 4998));
					System.out.println("WARN (GoogleSheetsUploader): cell value too long; value was cut at 5000 chars");
				}else {
//					if(v.startsWith("http://") || v.startsWith("https://")) {
//						recVals.add("=HYPERLINK(\"http://stackoverflow.com\",\"SO label\")"
//								
////								new CellData().setHyperlink(v)
////								.setUserEnteredValue(new ExtendedValue().setStringValue(v))
////								.setUserEnteredValue(new ExtendedValue()
////			                    .setFormulaValue("=HYPERLINK(\"http://stackoverflow.com\",\"SO label\")"))
////								.setFormulaValue("=HYPERLINK(\""+v+"\",\"link\")"))
//						);
//					}
					recVals.add(v);
				}
			}
			vals.add(recVals);
		}	
		parser.close();
		return GoogleSheetsApi.createValues(vals, cols);
	}

	public static String sheetTitleFromFileName(File csvFile) {
		String sheetTitle=csvFile.getName().substring(0, csvFile.getName().lastIndexOf('.'));
		return sheetTitle;
	}

}
