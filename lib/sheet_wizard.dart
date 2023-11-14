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
final List<List<List<String>>> schedules = [];

final List<int> nullableRows = [];
final List<String> files = [
  './assets/uteis.xlsx',
  // './assets/sabado.xlsx',
  // './assets/domingo.xlsx',
];
final List<File> excelFiles = [];

void start() {
  loadExcelFile();
}

//change this to your excel file location

void loadExcelFile() async {
  for (var i = 0; i < files.length; i++) {
    excelFiles.add(File(files[i]));
    print('done Loading files ${excelFiles[i].path}');
  }

  for (var i = 0; i < excelFiles.length; i++) {
    final bytes = excelFiles[i].readAsBytesSync();
    final response = await startConverting(bytes, excelFiles[i].path.replaceAll('xlsx', '').replaceAll('.', '').split('/').last);
    if (response != null) {
      print('done writing to file ${response.path}');
    } else {
      print('error writing to file');
    }
  }
}

/// Initiates a file with the given bytes.
///
/// [bytes] is a list of integers representing the bytes of the file.
///
/// FILEPATH: /C:/Users/atali/Desktop/My apps/sheet-wizard/lib/sheet_wizard.dart
Future startConverting(List<int> bytes, String fileName) async {
  final sheet = Excel.decodeBytes(bytes);
  //Você pode alterar o nome da table se quiser
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
  print(table.rows.first[21]?.value);
  int rowOffset = 0;
  for (var i = 0; i < table.rows.first.length; i++) {
    final element = table.rows.first[i];
    if (element?.value != null) {
      if (districts.contains(element!.value.toString())) {
        final index = districts.indexOf(element.value.toString());
        int nulableRowsCount = 0;
        int rowOffset = 0;

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
        for (var j = 2; j < rowsLength; j++) {
          final row = table.rows[j];
          List<String> values = [];
          dynamic valueOne;
          dynamic valueTwo;
          dynamic valueThree;
          dynamic valueFour;
          switch (nullableRows[index]) {
            case 1:
              if (index == 0) {
                valueOne = row[0]?.value ?? '';
                valueTwo = row[1]?.value ?? '';
              } else if (index > 0) {
                if (index == 2) {
                  rowOffset = 1;
                } else {
                  rowOffset = 3;
                }
                valueOne = row[0 + index * 2 + rowOffset]?.value ?? '';
                valueTwo = row[1 + index * 2 + rowOffset]?.value ?? '';
              }
              break;
            case 2:
              if (index == 0) {
                valueOne = row[0]?.value ?? '';
                valueTwo = row[1]?.value ?? '';
                valueThree = row[2]?.value ?? '';
              } else if (index > 0) {
                valueOne = row[0 + index > 0 ? index * 2 : index]?.value ?? '';
                valueTwo = row[1 + index > 0 ? index * 2 + 1 : index + 1]?.value ?? '';
                valueThree = row[2 + index > 0 ? index * 2 + 2 : index + 2]?.value ?? '';
              }
              break;
            case 3:
              rowOffset = 1;
              if (index > 3) {
                rowOffset = 3;
              } else {
                rowOffset = 1;
              }
              valueOne = row[0 + index * 2 + rowOffset]?.value ?? '';
              valueTwo = row[1 + index * 2 + rowOffset]?.value ?? '';
              valueThree = row[2 + index * 2 + rowOffset]?.value ?? '';
              valueFour = row[3 + index * 2 + rowOffset]?.value ?? '';
              break;
            default:
          }

          values.add(valueOne.toString());
          values.add(valueTwo.toString());
          values.add(valueThree.toString());
          values.add(valueFour.toString());

          schedules[index].add(values);
        }
      }
    }
  }

  schedules.removeWhere((element) => element.isEmpty);

  final Map<String, dynamic> schedulesJson = {};
  for (var i = 0; i < districts.length; i++) {
    districts[i] = districts[i].replaceAll('LINHA', '');
    schedulesJson[districts[i]] = schedules[i];
  }
  //remove the zeros from the nullableRows list
  final List<int> newNullables = nullableRows.sublist(0, 10);
  for (var i = 0; i < newNullables.length; i++) {
    newNullables[i] = newNullables[i] + 1;
  }
  // int selectedDistrict = 8;
  // print(districts[selectedDistrict]);
  // print(schedules[selectedDistrict]);
  // return;
  final response = await writeToFile('./assets/json/$fileName.json', jsonEncode(schedulesJson));
  if (response.path.isNotEmpty) {
    districts.clear();
    schedules.clear();
    nullableRows.clear();
    return response;
  } else {
    return null;
  }
}

Future<File> writeToFile(String fileName, String data) async {
  final file = File(fileName);
  return file.writeAsString(data).then((value) {
    print('File written');
    return value;
  });
}
