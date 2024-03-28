package apiclient.google.sheets;

import java.io.IOException;
import java.io.Reader;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.List;
import java.util.Map.Entry;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Get;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Clear;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Update;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.ValueRange;

import apiclient.google.GoogleApi;

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
	
	public void print(Object... values) {
		for(Object v: values)
			printVal(v);			
	}
	
	public void printRecord(Object... values) {
		for(Object v: values)
			printVal(v);			
		maxCols=Math.max(maxCols, line.size());
		line=new ArrayList<Object>();
		vals.add(line);
	}
	public void println() {
		maxCols=Math.max(maxCols, line.size());
		line=new ArrayList<Object>();
		vals.add(line);
	}

	private void printVal(Object obj) {
		line.add(obj==null ? "" : obj);
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
//			append.setValueInputOption("RAW");
			append.setValueInputOption("USER_ENTERED");
			append.execute();
		}
	}

	public String toCsv()  {
		StringBuffer sb=new StringBuffer();
		try {
			CSVPrinter printer=new CSVPrinter(sb, CSVFormat.DEFAULT);
			vals.forEach(cells -> {
				try {
					cells.forEach(cell -> {
						try {
							printer.print(cell.toString());
						} catch (IOException e) {
							//should  not hapen when writing to a StringBuffer
							throw new RuntimeException(e.getMessage(), e);
						}
					});
					printer.println();
				} catch (IOException e) {
					//should  not hapen when writing to a StringBuffer
					throw new RuntimeException(e.getMessage(), e);
				}
			});
			printer.close();
			return sb.toString();
		} catch (IOException e) {
			//should  not hapen when writing to a StringBuffer
			throw new RuntimeException(e.getMessage(), e);
		}		
	}

	public void printCsv(Reader reader) throws IOException {
		CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT);
		for(CSVRecord rec: csvParser) {
			for(String v: rec) {
				print(v);
			}
			println();
		}
		csvParser.close();
	}

	
}
