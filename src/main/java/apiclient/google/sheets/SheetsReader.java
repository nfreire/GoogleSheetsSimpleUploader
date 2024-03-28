package apiclient.google.sheets;

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

import apiclient.google.GoogleApi;

public class SheetsReader implements Iterable<List<Object>> {
	List<List<Object>> data;
	
	public SheetsReader(String spreadsheetId, String sheetTitle, String rangeDataToRead) throws IOException {
		Sheets service = GoogleApi.getSheetsService();
		data = service.spreadsheets().values()
	            .get(spreadsheetId, sheetTitle + "!" + rangeDataToRead)
	            .execute().getValues();
	}
	
	@Override
	public Iterator<List<Object>> iterator() {
		return data.iterator();
	}
}
