package com.googlesheet.component;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.googlesheet.GoogleAuthorizeUtil;
import com.googlesheet.model.Employee;

@Component
public class DataMigration {

	@Scheduled(fixedDelay = 100000)
	public void migrateData() throws Exception {
		List<Object> list = new ArrayList<>();
		for (int i = 0; i <= 5000; i++) {
			Employee e1 = new Employee();
			e1.setId("1");
			e1.setName("TestName1");
			e1.setSal("test sal5");
			e1.setCity(String.valueOf(new Date().getMinutes()));

			list.add(e1);
		}

		//update(list);
		append(list);
	}
	
	private void append(List<Object> list) throws Exception {
		Sheets service = getSheetsService();
		// Clear the sheet
		ClearValuesRequest clearrequestBody = new ClearValuesRequest();
		Sheets.Spreadsheets.Values.Clear request = service.spreadsheets().values()
				.clear("1PVTjkVPqLO2aQdYU0cVnWM-2SZ_tUiqNm_7NiKZQSkk", "A:AT", clearrequestBody);
		request.execute();
		
		ValueRange data = new ValueRange();
		
		List<List<Object>> values = new ArrayList<>();
		values.add(Arrays.asList("Employee Id1", "Employee Name", "Employee Sal", "Employee City", "Employee City",
				"Employee Id2", "Employee Name", "Employee Sal", "Employee City", "Employee City",
				"Employee Id3", "Employee Name", "Employee Sal", "Employee City", "Employee City",
				"Employee Id4", "Employee Name", "Employee Sal", "Employee City", "Employee City",
				"Employee Id5", "Employee Name", "Employee Sal", "Employee City", "Employee City",
				"Employee Id6", "Employee Name", "Employee Sal", "Employee City", "Employee City"));
		values.add(Arrays.asList("Employee Id", "Employee Name", "Employee Sal", "Employee City"));
		for (Object o : list) {
			Employee e = (Employee) o;
			values.add(Arrays.asList(String.valueOf(e.getId()), e.getName(), e.getSal(), e.getCity()));
		}
		
		data.setValues(values);
		
		AppendValuesResponse result =
		        service.spreadsheets().values().append("1PVTjkVPqLO2aQdYU0cVnWM-2SZ_tUiqNm_7NiKZQSkk", "A1", data)
		                .setValueInputOption("USER_ENTERED")
		                .execute();
		System.out.println(result);
	}

	/*private void update(List<Object> list) throws Exception {
		Sheets service = getSheetsService();
		// Clear the sheet
		ClearValuesRequest clearrequestBody = new ClearValuesRequest();
		Sheets.Spreadsheets.Values.Clear request = service.spreadsheets().values()
				.clear("1PVTjkVPqLO2aQdYU0cVnWM-2SZ_tUiqNm_7NiKZQSkk", "A:Z", clearrequestBody);
		request.execute();

		List<List<Object>> headerValues = new ArrayList<>();
		headerValues.add(Arrays.asList("Employee Id", "Employee Name", "Employee Sal", "Employee City"));

		ValueRange header = new ValueRange();
		header.setRange("A1:D1");
		header.setValues(headerValues);

		List<ValueRange> data = new ArrayList<>();
		data.add(header);

		int index = 1;
		for (Object o : list) {
			index++;
			String startRange = "A" + index;

			Employee e = (Employee) o;
			List<List<Object>> values = new ArrayList<>();
			values.add(Arrays.asList(e.getId()));
			values.add(Arrays.asList(e.getName()));
			values.add(Arrays.asList(e.getSal()));
			values.add(Arrays.asList(e.getCity()));

			ValueRange dataRange = new ValueRange();
			dataRange.setRange(startRange);
			dataRange.setValues(values);
			dataRange.setMajorDimension("COLUMNS");

			data.add(dataRange);
		}

		BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest();
		requestBody.setValueInputOption("RAW");
		requestBody.setData(data);

		Sheets.Spreadsheets.Values.BatchUpdate updateRequest = service.spreadsheets().values()
				.batchUpdate("1PVTjkVPqLO2aQdYU0cVnWM-2SZ_tUiqNm_7NiKZQSkk", requestBody);
		

		BatchUpdateValuesResponse response = updateRequest.execute();
		
		System.out.println(response);

	}*/

	public static Sheets getSheetsService() throws Exception {
		Credential credential = GoogleAuthorizeUtil.authorize();
		return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(),
				credential).setApplicationName("DataMigration").build();
	}
}