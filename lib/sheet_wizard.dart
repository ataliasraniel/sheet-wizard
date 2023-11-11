/// This file contains functions to load an excel file, extract data from it, and write the extracted data to a JSON file.
/// The extracted data is a list of schedules for different districts.
/// The file path for the excel file and the output JSON file are hardcoded.
/// The extracted data is written to the output JSON file in the form of a Map with district names as keys and their schedules as values.
/// The functions in this file are:
///   - start(): calls the loadExcelFile() function to initiate the loading of the excel file.
///   - loadExcelFile(): loads the excel file from the specified file path and calls the initiateFile() function to extract data from it.
///   - initiateFile(): extracts data from the loaded excel file and writes it to the output JSON file.
///   - writeToFile(): writes the extracted data to the output JSON file.
import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';

List<Map<String, dynamic>> busData = [];
List<String> districts = [];
final List<List<String>> schedules = [];

final List<int> nullableRows = [];

void start() {
  loadExcelFile();
}

//change this to your excel file location
final String excelFileLocationPath = './assets/uteis.xlsx';

void loadExcelFile() {
  final File excelFile = File(excelFileLocationPath);
  if (excelFile.existsSync()) {
    var bytes = excelFile.readAsBytesSync();
    initiateFile(bytes);
  } else {
    print('File does not exist');
  }
}

/// Initiates a file with the given bytes.
///
/// [bytes] is a list of integers representing the bytes of the file.
///
/// FILEPATH: /C:/Users/atali/Desktop/My apps/sheet-wizard/lib/sheet_wizard.dart
void initiateFile(List<int> bytes) {
  final sheet = Excel.decodeBytes(bytes);
  final Sheet table = sheet.tables['Table 1']!;
  for (var element in table.rows.first) {
    if (element != null && element.value != null) {
      districts.add(element.value.toString());
    }
  }

  final Map<String, dynamic> jsonData = {};
  for (var district in districts) {
    jsonData[district] = [];
  }
  final int rowsLength = table.rows.length;
  for (var i = 0; i < districts.length; i++) {
    schedules.add([]);
  }

  for (var element in table.rows.first) {
    if (element?.value != null) {
      if (districts.contains(element!.value.toString())) {
        final index = districts.indexOf(element.value.toString());
        print('iterating index $index');
        int nulableRowsCount = 0;
        for (var element in table.rows[0]) {
          //check if the element is null, and if it is, increment the nullable count
          //and when find a non null element, add the number to the list and reset the count and so on
          if (element == null || element.value == null) {
            nulableRowsCount++;
          } else {
            nullableRows.add(nulableRowsCount);
            nulableRowsCount = 0;
          }
        }
        nullableRows.removeWhere((element) => element == 0);

        final int actualIndex = nullableRows[index] + 1;
        for (var i = 3; i < rowsLength; i++) {
          final row = table.rows[i];
          List<String> values = [];
          final valueOne = row[0 + index > 0 ? index * 2 : index]?.value;
          final valueTwo = row[1 + index > 0 ? index * 2 + 1 : index * 2 + 1]?.value;

          if (valueOne == null || valueTwo == null) continue;
          values.add(valueOne.toString());
          values.add(valueTwo.toString());
          schedules[index].add(values.toString());
        }
      }
    }
  }

  schedules.removeWhere((element) => element.isEmpty);
  final Map<String, dynamic> schedulesJson = {};
  for (var i = 0; i < districts.length; i++) {
    schedulesJson[districts[i]] = schedules[i];
  }
  //remove the zeros from the nullableRows list
  final List<int> newNullables = nullableRows.sublist(0, 9);
  for (var i = 0; i < newNullables.length; i++) {
    newNullables[i] = newNullables[i] + 1;
  }
  // print(newNullables);

  writeToFile('./assets/uteis.json', jsonEncode(schedulesJson));
}

Future<File> writeToFile(String fileName, String data) async {
  final file = File(fileName);
  return file.writeAsString(data).then((value) {
    print('File written');
    return value;
  });
}

int _getCellLengthValue(Data? data) {
  return 0;
}
