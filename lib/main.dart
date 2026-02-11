import 'dart:io';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as ex;

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Auditor Pro',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        useMaterial3: true,
        colorSchemeSeed: Colors.indigo,
        scaffoldBackgroundColor: const Color(0xFFF5F7FA),
      ),
      home: const ExcelComparePage(),
    );
  }
}

class ExcelComparePage extends StatefulWidget {
  const ExcelComparePage({super.key});
  @override
  State<ExcelComparePage> createState() => _ExcelComparePageState();
}

class _ExcelComparePageState extends State<ExcelComparePage> {
  File? gstFile;
  File? aceFile;
  final sheetController = TextEditingController(text: 'B2B');
  final List<Map<String, dynamic>> results = [];
  final List<Map<String, String>> diffData = [];
  bool isProcessing = false;

  // Normalizes currency strings into clean decimals
  String _normalizeAmount(dynamic value) {
    if (value == null) return "0.00";
    String cleaned = value.toString().replaceAll(RegExp(r'[^\d.]'), '');
    if (cleaned.isEmpty || cleaned == ".") return "0.00";
    double? parsed = double.tryParse(cleaned);
    return parsed?.toStringAsFixed(2) ?? "0.00";
  }

  // Cleans invoice numbers for consistent matching
  String _cleanInv(dynamic value) {
    if (value == null) return "";
    return value.toString().trim().toUpperCase();
  }

  Future<File?> pickExcel() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xls', 'xlsx'],
    );
    if (result == null || result.files.single.path == null) return null;
    return File(result.files.single.path!);
  }

  /// DYNAMIC EXTRACTION: Finds headers regardless of row position
  List<Map<String, String>> _extractData(ex.Excel excel, List<String> targetSheetNames) {
    List<Map<String, String>> data = [];
    for (var name in targetSheetNames) {
      var sheet = excel.tables[name];
      if (sheet == null || sheet.maxRows < 1) continue;

      int headerIdx = -1;
      int idxInvNum = -1, idxTotalVal = -1, idxTaxVal = -1;

      // Scan up to 100 rows to find the header row (Handles GST starting at Row 4)
      int scanLimit = sheet.maxRows < 100 ? sheet.maxRows : 100;
      for (int i = 0; i < scanLimit; i++) {
        var row = sheet.rows[i];
        for (int j = 0; j < row.length; j++) {
          String val = row[j]?.value?.toString().toLowerCase().trim() ?? "";
          
          // Identify Invoice Column
          if ((val.contains('inv') || val.contains('note') || val.contains('bill')) && 
              (val.contains('no') || val.contains('num') || val.contains('number'))) {
            idxInvNum = j;
          }
          // Identify Total Amount Column
          if ((val.contains('val') || val.contains('amt') || val.contains('total')) && 
              (val.contains('inv') || val.contains('bill'))) {
            idxTotalVal = j;
          }
          // Identify Taxable Value Column
          if (val.contains('taxable')) {
            idxTaxVal = j;
          }
        }
        if (idxInvNum != -1) { 
          headerIdx = i; 
          break; 
        }
      }

      if (headerIdx == -1) continue;

      // Extract rows starting from AFTER the detected header
      for (int i = headerIdx + 1; i < sheet.rows.length; i++) {
        var row = sheet.rows[i];
        if (row.isEmpty || row.length <= idxInvNum) continue;

        String invNum = _cleanInv(row[idxInvNum]?.value);
        if (invNum.isEmpty || invNum == "TOTAL" || invNum == "GRAND TOTAL") continue;

        data.add({
          'InvNum': invNum,
          'TotalVal': _normalizeAmount(idxTotalVal != -1 && row.length > idxTotalVal ? row[idxTotalVal]?.value : '0.00'),
          'TaxVal': _normalizeAmount(idxTaxVal != -1 && row.length > idxTaxVal ? row[idxTaxVal]?.value : '0.00'),
          'Matched': 'false',
        });
      }
    }
    return data;
  }

  Future<void> compareSheets(ex.Excel gst, ex.Excel ace, String searchKey) async {
    setState(() {
      results.clear();
      diffData.clear();
      isProcessing = true;
    });

    final key = searchKey.trim().toLowerCase();
    final gstTargets = gst.tables.keys.where((s) => s.toLowerCase().contains(key)).toList();
    final aceTargets = ace.tables.keys.where((s) => s.toLowerCase().contains(key)).toList();

    List<Map<String, String>> gstRows = _extractData(gst, gstTargets);
    List<Map<String, String>> aceRows = _extractData(ace, aceTargets);

    if (gstRows.isEmpty || aceRows.isEmpty) {
      setState(() {
        results.add({'msg': "❌ Format Error: Sheet '$searchKey' not found or headers missing.", 'color': Colors.red});
        isProcessing = false;
      });
      return;
    }

    // Compare GST against ACE
    for (var gstRow in gstRows) {
      var match = aceRows.firstWhere(
        (ace) => ace['InvNum'] == gstRow['InvNum'] && 
                 ace['TotalVal'] == gstRow['TotalVal'] && 
                 ace['TaxVal'] == gstRow['TaxVal'] && 
                 ace['Matched'] == 'false',
        orElse: () => {},
      );

      if (match.isNotEmpty) {
        match['Matched'] = 'true';
        gstRow['Matched'] = 'true';
      } else {
        var partial = aceRows.where((ace) => ace['InvNum'] == gstRow['InvNum']).toList();
        
        double gstTax = double.tryParse(gstRow['TaxVal']!) ?? 0.0;
        double aceTax = partial.isNotEmpty ? (double.tryParse(partial.first['TaxVal']!) ?? 0.0) : 0.0;
        double taxDiff = (gstTax - aceTax).abs();
        String diffMsg = taxDiff > 0.01 ? " (${taxDiff.toStringAsFixed(2)} Tax Diff)" : "";

        diffData.add({
          'InvNum': gstRow['InvNum']!,
          'Status': partial.isNotEmpty ? 'Amount Mismatch' : 'Missing in ACE',
          'GST Total': gstRow['TotalVal']!, 'GST Taxable': gstRow['TaxVal']!,
          'ACE Total': partial.isNotEmpty ? partial.first['TotalVal']! : '0.00',
          'ACE Taxable': partial.isNotEmpty ? partial.first['TaxVal']! : '0.00',
          'Color': 'Red'
        });

        results.add({
          'msg': '❌ No: ${gstRow['InvNum']}: ${partial.isNotEmpty ? "Amount Mismatch$diffMsg" : "Not Found"}', 
          'color': Colors.red.shade800
        });
      }
    }

    // Check for extras in ACE
    for (var aceRow in aceRows) {
      if (aceRow['Matched'] == 'false') {
        diffData.add({
          'InvNum': aceRow['InvNum']!,
          'Status': 'Extra in ACE',
          'GST Total': '0.00', 'GST Taxable': '0.00',
          'ACE Total': aceRow['TotalVal']!, 'ACE Taxable': aceRow['TaxVal']!,
          'Color': 'Yellow'
        });
        results.add({
          'msg': '⚠️ No: ${aceRow['InvNum']} Extra in ACE (${aceRow['TotalVal']})', 
          'color': Colors.orange.shade900
        });
      }
    }
    setState(() { isProcessing = false; });
  }

  Future<void> exportExcel() async {
    if (diffData.isEmpty) return;
    String? path = await FilePicker.platform.saveFile(
      fileName: 'Audit_Report.xlsx', 
      allowedExtensions: ['xlsx'], 
      type: FileType.custom
    );
    if (path == null) return;

    var excel = ex.Excel.createExcel();
    ex.Sheet sheet = excel['Audit'];
    var headerStyle = ex.CellStyle(bold: true, backgroundColorHex: ex.ExcelColor.fromHexString('#CCCCCC'));
    
    List<String> h = ['No/Inv Number', 'Status', 'GST Total', 'GST Tax', 'ACE Total', 'ACE Tax'];
    for (int i = 0; i < h.length; i++) {
      var cell = sheet.cell(ex.CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0));
      cell.value = ex.TextCellValue(h[i]);
      cell.cellStyle = headerStyle;
    }

    for (int i = 0; i < diffData.length; i++) {
      var row = diffData[i];
      var style = ex.CellStyle(backgroundColorHex: ex.ExcelColor.fromHexString(row['Color'] == 'Red' ? '#FFCCCC' : '#FFFFCC'));
      sheet.cell(ex.CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i+1)).value = ex.TextCellValue(row['InvNum']!);
      sheet.cell(ex.CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i+1)).value = ex.TextCellValue(row['Status']!);
      
      // Numbers are saved as DoubleCellValue to allow "Format Painting" in Excel
      _cell(sheet, 2, i+1, row['GST Total']!, style);
      _cell(sheet, 3, i+1, row['GST Taxable']!, style);
      _cell(sheet, 4, i+1, row['ACE Total']!, style);
      _cell(sheet, 5, i+1, row['ACE Taxable']!, style);
    }

    final bytes = excel.encode();
    if (bytes != null) {
      try {
        await File(path).writeAsBytes(bytes);
        if (mounted) ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content: Text("✅ Saved!"), backgroundColor: Colors.green));
      } catch (e) {
        if (mounted) {
          showDialog(
            context: context,
            builder: (ctx) => AlertDialog(
              title: const Text("⚠️ File Locked"),
              content: const Text("Close 'Audit_Report.xlsx' in Excel and try again."),
              actions: [TextButton(onPressed: () => Navigator.pop(ctx), child: const Text("OK"))],
            ),
          );
        }
      }
    }
  }

  void _cell(ex.Sheet s, int c, int r, String v, ex.CellStyle st) {
    var cell = s.cell(ex.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r));
    cell.value = ex.DoubleCellValue(double.tryParse(v) ?? 0.00);
    cell.cellStyle = st;
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Center(
        child: Container(
          constraints: const BoxConstraints(maxWidth: 900),
          margin: const EdgeInsets.all(40),
          decoration: BoxDecoration(
            color: Colors.white,
            borderRadius: BorderRadius.circular(24),
            boxShadow: [BoxShadow(color: Colors.black.withOpacity(0.05), blurRadius: 20, offset: const Offset(0, 10))],
          ),
          child: Column(
            children: [
              Padding(
                padding: const EdgeInsets.fromLTRB(30, 30, 30, 10),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    const Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text("Auditor Pro", style: TextStyle(fontSize: 24, fontWeight: FontWeight.w900, color: Colors.indigo)),
                        Text("GST vs ACE Smart Comparison", style: TextStyle(fontSize: 12, color: Colors.grey)),
                      ],
                    ),
                    if (diffData.isNotEmpty)
                      ElevatedButton.icon(
                        onPressed: exportExcel, 
                        icon: const Icon(Icons.download),
                        label: const Text("Export Report"),
                      )
                  ],
                ),
              ),
              const Divider(indent: 30, endIndent: 30),
              Padding(
                padding: const EdgeInsets.symmetric(horizontal: 30, vertical: 15),
                child: Row(
                  children: [
                    Expanded(
                      child: TextField(
                        controller: sheetController,
                        decoration: InputDecoration(
                          prefixIcon: const Icon(Icons.table_chart_outlined, size: 20),
                          hintText: "Target Sheet Name (e.g., B2B)",
                          filled: true,
                          fillColor: const Color(0xFFF1F3F4),
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(12), borderSide: BorderSide.none),
                        ),
                      ),
                    ),
                    const SizedBox(width: 12),
                    ElevatedButton(
                      style: ElevatedButton.styleFrom(
                        backgroundColor: Colors.indigo,
                        foregroundColor: Colors.white,
                        padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 20),
                        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                      ),
                      onPressed: (gstFile != null && aceFile != null && !isProcessing)
                        ? () async {
                            showDialog(
                              context: context,
                              barrierDismissible: false,
                              builder: (ctx) => const AlertDialog(
                                content: Column(
                                  mainAxisSize: MainAxisSize.min,
                                  children: [
                                    CircularProgressIndicator(),
                                    SizedBox(height: 20),
                                    Text("Analyzing dynamic layouts...")
                                  ],
                                ),
                              ),
                            );
                            try {
                              final g = ex.Excel.decodeBytes(await gstFile!.readAsBytes());
                              final a = ex.Excel.decodeBytes(await aceFile!.readAsBytes());
                              await compareSheets(g, a, sheetController.text);
                            } finally {
                              if (mounted) Navigator.pop(context);
                            }
                          }
                        : null,
                      child: const Text("Run Audit"),
                    ),
                  ],
                ),
              ),
              Padding(
                padding: const EdgeInsets.symmetric(horizontal: 30),
                child: Row(
                  children: [
                    _fileCard("GST Source", gstFile, () async { gstFile = await pickExcel(); setState(() {}); }),
                    const SizedBox(width: 12),
                    _fileCard("ACE Source", aceFile, () async { aceFile = await pickExcel(); setState(() {}); }),
                  ],
                ),
              ),
              Expanded(
                child: Container(
                  margin: const EdgeInsets.all(30),
                  decoration: BoxDecoration(
                    color: const Color(0xFFFAFBFC),
                    borderRadius: BorderRadius.circular(16),
                    border: Border.all(color: const Color(0xFFE0E3E7)),
                  ),
                  child: results.isEmpty 
                    ? const Center(child: Text("Ready for audit. Upload files to start."))
                    : ListView.builder(
                        padding: const EdgeInsets.all(10),
                        itemCount: results.length,
                        itemBuilder: (c, i) => Card(
                          elevation: 0,
                          color: results[i]['color'].withOpacity(0.08),
                          child: ListTile(
                            leading: Icon(results[i]['msg'].contains('❌') ? Icons.error_outline : Icons.warning_amber_rounded, color: results[i]['color']),
                            title: Text(results[i]['msg'], style: TextStyle(color: results[i]['color'], fontSize: 13, fontWeight: FontWeight.bold)),
                          ),
                        ),
                      ),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Widget _fileCard(String title, File? file, VoidCallback tap) {
    bool hasFile = file != null;
    return Expanded(
      child: InkWell(
        onTap: tap,
        borderRadius: BorderRadius.circular(12),
        child: Container(
          padding: const EdgeInsets.all(16),
          decoration: BoxDecoration(
            border: Border.all(color: hasFile ? Colors.green : const Color(0xFFE0E3E7)),
            borderRadius: BorderRadius.circular(12),
            color: hasFile ? Colors.green.withOpacity(0.05) : Colors.transparent,
          ),
          child: Row(
            children: [
              Icon(hasFile ? Icons.check_circle : Icons.upload_file, color: hasFile ? Colors.green : Colors.indigo),
              const SizedBox(width: 10),
              Flexible(child: Text(hasFile ? file.path.split(Platform.pathSeparator).last : title, overflow: TextOverflow.ellipsis)),
            ],
          ),
        ),
      ),
    );
  }
}