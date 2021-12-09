package inescid.util.googlesheets;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map.Entry;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Get;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Clear;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Update;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.ValueRange;

public class SheetsPrinter {
	String spreadsheetId;
	String sheetTitle;
	
	List<List<Object>> vals=new ArrayList<List<Object>>();
	int maxCols=0;
//	int col=0;
	ArrayList<Object> line;

	public SheetsPrinter(String spreadsheetId, String sheetTitle) {
		super();
		this.spreadsheetId = spreadsheetId;
		this.sheetTitle = sheetTitle;
		line=new ArrayList<Object>();
		vals.add(line);
	}
	
	public void print(Object value) {
		line.add(value);
	}
	
	public void printRecord(Object... values) {
		for(Object v: values)
			line.add(v);			
		maxCols=Math.max(maxCols, line.size());
		line=new ArrayList<Object>();
		vals.add(line);
	}
	public void println() {
		maxCols=Math.max(maxCols, line.size());
		line=new ArrayList<Object>();
		vals.add(line);
	}
	
	public void close() throws IOException {
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
		if(!sheetExists) 
			GoogleSheetsApi.addSheet(spreadsheetId, sheetTitle);
		 
		if(!vals.isEmpty() && maxCols>0) {
			Entry<ValueRange, Integer> createValues = GoogleSheetsApi.createValues(vals, maxCols);
			ValueRange vRange=createValues.getKey();
			int cols=createValues.getValue();
			int rows=vRange.getValues().size();
			
			Update append = service.spreadsheets().values().update(spreadsheetId, GoogleSheetsApi.makeRangeExpression(sheetTitle, rows, cols) , vRange);
			append.setValueInputOption("RAW");
			append.execute();
		}
	}

	
}
