import 'dart:developer';
import 'dart:io';
import 'dart:typed_data';

import 'package:file_saver/file_saver.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as excel;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatefulWidget {
  const MyApp({Key? key}) : super(key: key);
  @override
  _MyAppState createState() => _MyAppState();
}

class _MyAppState extends State<MyApp> {
  @override
  void initState() {
    super.initState();
  }

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: Scaffold(
        appBar: AppBar(
          title: const Text('Export Excel'),
        ),
        body: Center(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              if (!kIsWeb)
                if (Platform.isAndroid || Platform.isIOS || Platform.isMacOS)
                  ElevatedButton(
                    onPressed: () async {
                      generateExcel();
                    },
                    child: const Text("Download Excel"),
                  )
            ],
          ),
        ),
      ),
    );
  }

  Future<void> generateExcel() async {
    //Create a Excel document.

    //Creating a workbook.
    final excel.Workbook workbook = excel.Workbook();
    //Accessing via index
    final excel.Worksheet sheet = workbook.worksheets[0];
    sheet.showGridlines = false;

    // Enable calculation for worksheet.
    sheet.enableSheetCalculations();

    //Set data in the worksheet.
    sheet.getRangeByName('A1').value = "Nama";
    sheet.getRangeByName('B1').value = "Jenis Kelamin";
    sheet.getRangeByName('C1').value = "Telepon";
    sheet.getRangeByName('D1').value = "Tanggal Lahir";
    sheet.getRangeByName('E1').value = "Alamat";

    for (int i = 2; i < 10; i++) {
      sheet.getRangeByName('A$i').value = "$i Budi";
      sheet.getRangeByName('B$i').value = "$i Laki-laki";
      sheet.getRangeByName('C$i').value = "$i 0812xxx";
      sheet.getRangeByName('D$i').value = "$i 1992-01-01";
      sheet.getRangeByName('E$i').value = "$i Jl Rumah";
    }

    //Save and launch the excel.
    final List<int> sheets = workbook.saveAsStream();
    //Dispose the document.
    workbook.dispose();

    //Save and launch the file.
    await saveAndLaunchFile(sheets, 'namafile');
  }

  saveAndLaunchFile(sheets, name) async {
    Uint8List data = Uint8List.fromList(sheets);
    MimeType type = MimeType.MICROSOFTEXCEL;
    String path = await FileSaver.instance.saveAs("$name", data, "xlsx", type);
    log(path);
  }
}
