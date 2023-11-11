import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';

List<Map<String, dynamic>> busData = [];
List<String> districts = [];
final List<List<String>> schedules = [];

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
  print(schedulesJson);

  writeToFile('./assets/uteis.json', jsonEncode(schedulesJson));
}

Future<File> writeToFile(String fileName, String data) async {
  final file = File(fileName);
  return file.writeAsString(data).then((value) {
    print('File written');
    return value;
  });
}
