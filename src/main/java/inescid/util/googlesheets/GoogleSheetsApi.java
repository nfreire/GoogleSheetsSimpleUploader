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

public class GoogleSheetsApi {

	public static String create(String spreadsheetTitle, String sheetTitle) throws IOException {
		Sheets sheetsService = GoogleApi.getSheetsService();
		Spreadsheet spreadsheet = null;
		{
			Spreadsheet requestBody = new Spreadsheet();
			SpreadsheetProperties properties = new SpreadsheetProperties();
			properties.setTitle(spreadsheetTitle);
			requestBody.setProperties(properties);
			Sheets.Spreadsheets.Create request = sheetsService.spreadsheets().create(requestBody);
			spreadsheet = request.execute();
		}
		{
			Drive driveService = GoogleApi.getDriveService();
	        Permission permission = new Permission();
	        permission.setType("anyone");
	        permission.setRole("commenter");
			Permissions.Create permissionsRequest=driveService.permissions().create(spreadsheet.getSpreadsheetId(), permission);
			permissionsRequest.execute();
		}
		{
			List<Request> requests = new ArrayList<>();
			// Change the spreadsheet's title.
			requests.add(new Request()
			        .setUpdateSheetProperties(new UpdateSheetPropertiesRequest()
			                .setProperties(new SheetProperties()
			                        .setTitle(sheetTitle))
			                .setFields("title")));
			
			BatchUpdateSpreadsheetRequest body =
			        new BatchUpdateSpreadsheetRequest().setRequests(requests);
			        sheetsService.spreadsheets().batchUpdate(spreadsheet.getSpreadsheetId(), body).execute();
		}
		return spreadsheet.getSpreadsheetId();
	}
	public static void addSheet(String spreadsheetId, String sheetTitle) throws IOException {
		Sheets sheetsService = GoogleApi.getSheetsService();
		List<Request> requests = new ArrayList<>();
		// Change the spreadsheet's title.
		AddSheetRequest addSheetRequest = new AddSheetRequest();
		addSheetRequest.setProperties(new SheetProperties());
		addSheetRequest.getProperties().setTitle(sheetTitle);
		requests.add(new Request()
		        .setAddSheet(addSheetRequest));
		
		BatchUpdateSpreadsheetRequest body =
		        new BatchUpdateSpreadsheetRequest().setRequests(requests);
		        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
	}

	
	static Entry<ValueRange, Integer> createValues(List<List<Object>> vals, int maxCols) throws IOException {
		ValueRange vRange=new ValueRange();
		vRange.setValues(vals);
		return new DefaultMapEntry<ValueRange, Integer>(vRange,maxCols);
	}
	
	static String makeRangeExpression(String sheetTitle, int rows, int cols) {
		String lastCol;
		if(cols>26) {
			int lastLetter=(int)'A'+(cols % 26)-1;
			int firstLetter=(int)'A'+(cols / 26)-1;
			lastCol=""+(char)firstLetter+(char)lastLetter;			
		} else
			lastCol=""+(char)((int)'A'+cols-1);
		String range=(sheetTitle!=null ? sheetTitle+"!" : "")+"A1:"+lastCol+rows;
		System.out.println("rows: "+rows+ " ; maxCols: "+cols);
		return range;
	}

}
