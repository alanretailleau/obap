import 'dart:developer';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'dart:io';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: ExcelProcessor(),
    );
  }
}

// Obtenir la clé à partir de la colonne
String getKeyFromColumn(int col, int row) {
  switch (col) {
    case 2:
      switch (row) {
        case 3:
        case 13:
        case 23:
        case 32:
        case 42:
        case 52:
        case 62:
          return 'CMS';
        case 4:
        case 14:
        case 24:
        case 33:
        case 43:
        case 53:
        case 63:
          return 'PARC(CMS)';
        case 5:
        case 15:
        case 25:
        case 34:
        case 44:
        case 54:
        case 64:
          return 'CPE(CMS)';
        case 6:
        case 16:
        case 26:
        case 35:
        case 45:
        case 55:
        case 65:
          return 'SEC0';
        case 7:
        case 17:
        case 27:
        case 36:
        case 46:
        case 56:
        case 66:
          return 'PRE0';
        case 8:
        case 18:
        case 28:
        case 37:
        case 47:
        case 57:
        case 67:
          return 'PAR0';
        case 11:
        case 21:
        case 30:
        case 40:
        case 50:
        case 60:
        case 70:
          return 'MRA0';
        case 12:
        case 22:
        case 31:
        case 41:
        case 51:
        case 61:
        case 71:
          return 'CMSR';
        default:
          return '';
      }
    case 4:
      switch (row) {
        case 3:
        case 13:
        case 23:
        case 32:
        case 42:
        case 52:
        case 62:
          return 'RP';
        case 4:
        case 14:
        case 24:
        case 33:
        case 43:
        case 53:
        case 63:
          return 'PARC(RP)';
        case 5:
        case 15:
        case 25:
        case 34:
        case 44:
        case 54:
        case 64:
          return 'CPE(RP)';
        case 6:
        case 16:
        case 26:
        case 35:
        case 45:
        case 55:
        case 65:
          return 'SEC1';
        case 7:
        case 17:
        case 27:
        case 36:
        case 46:
        case 56:
        case 66:
          return 'PRE1';
        case 8:
        case 18:
        case 28:
        case 37:
        case 47:
        case 57:
        case 67:
          return 'PAR1';
        case 9:
        case 19:
        case 29:
        case 38:
        case 48:
        case 58:
        case 68:
          return 'EPR1';
        case 11:
        case 21:
        case 30:
        case 40:
        case 50:
        case 60:
        case 70:
          return 'MRA1';
        case 12:
        case 22:
        case 31:
        case 41:
        case 51:
        case 61:
        case 71:
          return 'RPR';
        default:
          return '';
      }
    case 3:
      switch (row) {
        case 3:
        case 13:
        case 23:
        case 32:
        case 42:
        case 52:
        case 62:
          return 'IP';
        case 4:
        case 14:
        case 24:
        case 33:
        case 43:
        case 53:
        case 63:
          return 'PARC(IP)';
        case 5:
        case 15:
        case 25:
        case 34:
        case 44:
        case 54:
        case 64:
          return 'CPE(IP)';
        case 6:
        case 16:
        case 26:
        case 35:
        case 45:
        case 55:
        case 65:
          return 'SEC2';
        case 7:
        case 17:
        case 27:
        case 36:
        case 46:
        case 56:
        case 66:
          return 'PRE2';
        case 8:
        case 18:
        case 28:
        case 37:
        case 47:
        case 57:
        case 67:
          return 'PAR2';
        case 11:
        case 21:
        case 30:
        case 40:
        case 50:
        case 60:
        case 70:
          return 'MRA2';
        case 12:
        case 22:
        case 31:
        case 41:
        case 51:
        case 61:
        case 71:
          return 'IPR';
        default:
          return '';
      }
    default:
      return '';
  }
}

class ExcelProcessor extends StatefulWidget {
  @override
  _ExcelProcessorState createState() => _ExcelProcessorState();
}

class _ExcelProcessorState extends State<ExcelProcessor> {
  List<String?> _filePaths = [];
  int _year = 2023;
  String? _outputDirectory;

  final List<String> typologies = [
    'GV',
    'Autoclave CAFR',
    'RECIPIENTS FIXES',
    'Récipients à pression simple RPS',
    'SF-CTP Groupe froid selon CTP',
    'Tuyauterie',
    'Autre équipement T7 à rajouter ulterieurement'
  ];

  final Map<String, List<int>> typologyRanges = {
    'GV': [3, 12],
    'Autoclave CAFR': [13, 22],
    'RECIPIENTS FIXES': [23, 31],
    'Récipients à pression simple RPS': [32, 41],
    'SF-CTP Groupe froid selon CTP': [42, 51],
    'Tuyauterie': [52, 61],
    'Autre équipement T7 à rajouter ulterieurement': [62, 71]
  };

  final List<int> excludeCMS = [
    9,
    10,
    19,
    20,
    29,
    38,
    39,
    48,
    49,
    58,
    59,
    68,
    69
  ];
  final List<int> excludeRP = [10, 20, 30, 39, 49, 59, 69];
  final List<int> excludeIP = [
    9,
    10,
    19,
    20,
    29,
    38,
    39,
    48,
    49,
    58,
    59,
    68,
    69
  ];

  Map<String, List<List<int>>> generateRowIndexMap(int start, int end) {
    List<List<int>> generateIndices(List<int> excludeList, int column) {
      return List.generate(end - start + 1, (index) => index + start)
          .where((index) => !excludeList.contains(index))
          .map((index) => [index, column])
          .toList();
    }

    return {
      'CMS': generateIndices(excludeCMS, 2),
      'RP': generateIndices(excludeRP, 4),
      'IP': generateIndices(excludeIP, 3),
    };
  }

  Map<String, Map<String, List<List<int>>>> rowIndexMap = {};

  @override
  void initState() {
    super.initState();
    typologyRanges.forEach((typology, range) {
      rowIndexMap[typology] = generateRowIndexMap(range[0], range[1]);
    });
  }

  Future<void> _pickFiles() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
      allowMultiple: true,
    );

    if (result != null) {
      setState(() {
        _filePaths = result.paths;
      });
    }
  }

  Future<void> _pickOutputDirectory() async {
    String? selectedDirectory = await FilePicker.platform.getDirectoryPath();

    if (selectedDirectory != null) {
      setState(() {
        _outputDirectory = selectedDirectory;
      });
    }
  }

  Future<void> _processFiles() async {
    if (_filePaths.isEmpty || _outputDirectory == null) return;

    // Créer un nouveau fichier Excel
    var newExcel = Excel.createExcel();
    var bddSheetNew = newExcel['BDD_rex_confidentiel'];

    // Ajouter les en-têtes
    List<String> headers = [
      'AC',
      'Num contri',
      'ABREV',
      'TYP',
      'API/SPI',
      'CMS',
      'PARC(CMS)',
      'CPE(CMS)',
      'SEC0',
      'PRE0',
      'PAR0',
      'MRA0',
      'CMSR',
      'RP',
      'PARC(RP)',
      'CPE(RP)',
      'SEC1',
      'PRE1',
      'PAR1',
      'EPR1',
      'MRA1',
      'RPR',
      'IP',
      'PARC(IP)',
      'CPE(IP)',
      'SEC2',
      'PRE2',
      'PAR2',
      'MRA2',
      'IPR',
      'Commentaires et Analyse'
    ];
    bddSheetNew
        .appendRow(headers.map((header) => TextCellValue(header)).toList());

    for (String? filePath in _filePaths) {
      if (filePath == null) continue;

      var file = File(filePath);
      var bytes = file.readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      // Lire les feuilles
      var withPi = excel[' Avec PI'];
      var withoutPi = excel['Sans PI'];

      // Fonction pour extraire les données
      Map<String, String> extractData(Sheet sheet, List<List<int>> indices) {
        Map<String, String> data = {};
        indices.forEach((index) {
          int row = index[0];
          int col = index[1];
          String key = getKeyFromColumn(col, row);
          if (row < sheet.maxRows && col < sheet.maxColumns) {
            data[key] = sheet.row(row)[col]?.value?.toString() ?? "";
          } else if (data[key] == null) {
            data[key] = "";
          }
        });
        return data;
      }

      // Mapper et ajouter les données à la feuille
      void mapDataToRowAndAppend(Map<String, String> data, String typology,
          String apiSpi, String contriNum) {
        List<CellValue> row = [
          TextCellValue(_year.toString()), // AC
          TextCellValue(contriNum), // Num contri
          TextCellValue(''), // ABREV
          TextCellValue(typology), // TYP
          TextCellValue(apiSpi), // API/SPI
          TextCellValue(data['CMS'] ?? ""), // CMS
          TextCellValue(data['PARC(CMS)'] ?? ""), // PARC(CMS)
          TextCellValue(data['CPE(CMS)'] ?? ""), // CPE(CMS)
          TextCellValue(data['SEC0'] ?? ""), // SEC0
          TextCellValue(data['PRE0'] ?? ""), // PRE0
          TextCellValue(data['PAR0'] ?? ""), // PAR0
          TextCellValue(data['MRA0'] ?? ""), // MRA0
          TextCellValue(data['CMSR'] ?? ""), // CMSR
          TextCellValue(data['RP'] ?? ""), // RP
          TextCellValue(data['PARC(RP)'] ?? ""), // PARC(RP)
          TextCellValue(data['CPE(RP)'] ?? ""), // CPE(RP)
          TextCellValue(data['SEC1'] ?? ""), // SEC1
          TextCellValue(data['PRE1'] ?? ""), // PRE1
          TextCellValue(data['PAR1'] ?? ""), // PAR1
          TextCellValue(data['EPR1'] ?? ""), // EPR1
          TextCellValue(data['MRA1'] ?? ""), // MRA1
          TextCellValue(data['RPR'] ?? ""), // RPR
          TextCellValue(data['IP'] ?? ""), // IP
          TextCellValue(data['PARC(IP)'] ?? ""), // PARC(IP)
          TextCellValue(data['CPE(IP)'] ?? ""), // CPE(IP)
          TextCellValue(data['SEC2'] ?? ""), // SEC2
          TextCellValue(data['PRE2'] ?? ""), // PRE2
          TextCellValue(data['PAR2'] ?? ""), // PAR2
          TextCellValue(data['MRA2'] ?? ""), // MRA2
          TextCellValue(data['IPR'] ?? ""), // IPR
          TextCellValue(
              data['Commentaires et Analyse'] ?? ""), // Commentaires et Analyse
        ];
        bddSheetNew.appendRow(row);
      }

      // Traiter chaque typologie
      void processTypology(String typology) {
        String contriNum = filePath.split('/').last.split('-').first;
        rowIndexMap[typology]?.forEach((section, indices) {
          var data = extractData(withoutPi, indices);
          mapDataToRowAndAppend(data, typology, 'SPI', contriNum);
          var data2 = extractData(withPi, indices);
          mapDataToRowAndAppend(data2, typology, 'API', contriNum);
        });
      }

      typologies.forEach((typology) {
        processTypology(typology);
      });
    }

    // Sauvegarder les modifications dans un fichier Excel dans le répertoire choisi
    String outputPath =
        '$_outputDirectory/23-Fichier REX collecte controles 2023_updated.xlsx';
    var outputFile = File(outputPath);
    await outputFile.writeAsBytes(newExcel.encode()!);

    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text('Fichier traité et sauvegardé à $outputPath')),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Excel Processor'),
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          children: <Widget>[
            ElevatedButton(
              onPressed: _pickFiles,
              child: Text('Sélectionner des fichiers Excel'),
            ),
            if (_filePaths.isNotEmpty)
              Text("${_filePaths.length} fichiers sélectionnés"),
            TextField(
              decoration: InputDecoration(labelText: 'Année de collecte'),
              keyboardType: TextInputType.number,
              onChanged: (value) {
                setState(() {
                  _year = int.tryParse(value) ?? 2023;
                });
              },
            ),
            SizedBox(height: 20),
            ElevatedButton(
              onPressed: () async {
                await _pickOutputDirectory();
                if (_outputDirectory != null) {
                  ScaffoldMessenger.of(context).showSnackBar(
                    SnackBar(
                        content: Text(
                            'Répertoire de sortie sélectionné: $_outputDirectory')),
                  );
                  _processFiles();
                }
              },
              child: Text('Traiter les fichiers'),
            ),
          ],
        ),
      ),
    );
  }
}
