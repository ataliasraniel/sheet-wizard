import 'dart:convert';
import 'dart:developer';
import 'dart:io';

import 'package:excel/excel.dart';

List<Map<String, dynamic>> busData = [];

void start() {
  loadExcelFile();
}

//change this to your excel file location
final String excelFileLocationPath = './assets/carta.xlsx';

void loadExcelFile() {
  final File excelFile = File(excelFileLocationPath);
  if (excelFile.existsSync()) {
    print('File exists');
    var bytes = excelFile.readAsBytesSync();
    initiateFile(bytes);
  } else {
    print('File does not exist');
  }
}

void initiateFile(bytes) {
  var excel = Excel.decodeBytes(bytes);
  for (var i = 0; i < excel.tables.entries.length; i++) {
    getOneTable(excel.tables.entries.toList()[i].value);
  }
  serializeToJson();
}

void serializeToJson() {
  writeToFile('./busschedule.json', json.encode(busData));
}

void getOneTable(Sheet sheet) {
  List<String> leftTime = [];
  List<String> arriveTime = [];
  for (var i = 4; i < sheet.rows.length; i++) {
    leftTime.add(sheet.rows[i][2]?.value.toString() ?? '--');
  }
  for (var i = 4; i < sheet.rows.length; i++) {
    arriveTime.add(sheet.rows[i][3]?.value.toString() ?? '--');
  }
  List<Map<String, dynamic>> schedule = [];
  for (var i = 0; i < leftTime.length; i++) {
    schedule.add({
      'leftTime': leftTime[i],
      'arriveTime': arriveTime[i],
    });
  }
  Map<String, dynamic> data = {
    'header': sheet.sheetName,
    'schedule': schedule,
  };
  busData.add(data);
}

Future<File> writeToFile(String fileName, String data) async {
  final file = File(fileName);
  return file.writeAsString(data).then((value) {
    print('File written');
    return value;
  });
}
