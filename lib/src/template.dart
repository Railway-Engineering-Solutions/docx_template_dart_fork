import 'package:archive/archive.dart';
import 'package:collection/collection.dart' show IterableExtension;
import 'package:docx_template/docx_template.dart';
import 'package:docx_template/src/tag_models.dart';
import 'package:docx_template/src/view_manager.dart';
import 'package:xml/xml.dart';

import 'docx_entry.dart';
import 'editor.dart';

class DocxTemplateException implements Exception {
  final String message;

  DocxTemplateException(this.message);

  @override
  String toString() => message;
}

///
/// Sdt tags policy enum
///
/// [removeAll] - remove all sdt tags from document
///
/// [saveNullified] - save ONLY tags where [Content] is null
///
/// [saveText] - save ALL TextContent field (include nullifed [Content])
///
enum TagPolicy { removeAll, saveNullified, saveText }

///
/// Image save policy
///
/// [remove] - remove template image from generated document if [ImageContent] is null
///
/// [save] - save template image in generated document if [ImageContent] is null
///
enum ImagePolicy { remove, save }

class DocxTemplate {
  DocxTemplate._();
  late DocxManager _manager;

  ///
  /// Load Template from byte buffer of docx file
  ///
  static Future<DocxTemplate> fromBytes(List<int> bytes) async {
    final component = DocxTemplate._();
    final arch = ZipDecoder().decodeBytes(bytes, verify: true);

    component._manager = DocxManager(arch);

    return component;
  }

//   exportPdf() async {
//     var configuration = Configuration('9849d3fc-3eb2-442a-a085-8d21d92c3ad3',
//         '798d958e76c462d62b41be3d754a9d25');
//     var wordsApi = WordsApi(configuration);
// // Upload file to cloud
//     var localFileContent = await (File('generated.docx').readAsBytes());
//     var uploadRequest = UploadFileRequest(
//         ByteData.view(localFileContent.buffer), 'fileStoredInCloud.docx');
//     await wordsApi.uploadFile(uploadRequest);
//
// // Save file as pdf in cloud
//     var saveOptionsData = PdfSaveOptionsData()
//       ..fileName = 'destStoredInCloud.pdf';
//     var saveAsRequest =
//         SaveAsRequest('fileStoredInCloud.docx', saveOptionsData);
//     await wordsApi.saveAs(saveAsRequest);
//   }

  ///
  /// Get all tags with enhanced metadata including type and nesting information
  ///
  /// This method parses the DOCX XML structure to:
  /// - Identify tag types by examining View types
  /// - Detect nested tags within tables
  /// - Track parent-child relationships
  ///
  /// Returns a DocxTagCollection with comprehensive tag information.
  DocxTagCollection getTagsEnhanced() {
    final viewManager = ViewManager.attach(
      DocxManager(_manager.arch),
    );

    final allTags = <DocxTag>[];
    final tagsByType = <TagType, List<DocxTag>>{};
    final tableTags = <String, List<DocxTag>>{};
    final documentTags = <DocxTag>[];

    // Helper function to determine tag type from View
    TagType _getTagTypeFromView(View view) {
      if (view is RowView) {
        return TagType.table;
      } else if (view is ImgView) {
        return TagType.image;
      } else if (view is ListView) {
        return TagType.list;
      } else if (view is PlainView) {
        return TagType.plain;
      } else if (view is TextView) {
        // Check if it's a link by examining the original tag
        if (view.sdtView?.tag == 'link') {
          return TagType.link;
        }
        // Check for form field types in SDT properties
        final sdtPr = view.sdtView?.sdt.findElements('sdtPr').firstOrNull;
        if (sdtPr != null) {
          // Check for combobox
          if (sdtPr.findElements('dropDownList').isNotEmpty) {
            return TagType.combobox;
          }
          // Check for checkbox
          if (sdtPr.findElements('checkbox').isNotEmpty) {
            return TagType.checkbox;
          }
          // Check for date
          if (sdtPr.findElements('date').isNotEmpty) {
            return TagType.date;
          }
          // Check for rich text
          if (sdtPr.findElements('richText').isNotEmpty) {
            return TagType.richText;
          }
        }
        return TagType.text;
      }
      return TagType.text;
    }

    // Helper function to check if view is nested in a table
    RowView? _findParentRowView(View view) {
      View? current = view.parentView;
      while (current != null) {
        if (current is RowView) {
          return current;
        }
        current = current.parentView;
      }
      return null;
    }

    // Helper function to calculate row and column indices from XML
    Map<String, int?> _calculateTableIndices(View view) {
      final sdtElement = view.sdtView?.sdt;
      if (sdtElement == null) return {'row': null, 'column': null};

      // Find the containing table cell (w:tc)
      XmlElement? current = sdtElement;
      XmlElement? tableCell;
      XmlElement? tableRow;
      XmlElement? table;

      // Traverse up to find table structure
      while (current != null) {
        final namespaceUri = current.name.namespaceUri;
        if (current.name.local == 'tc' &&
            namespaceUri != null &&
            namespaceUri.contains('word')) {
          tableCell = current;
        } else if (current.name.local == 'tr' &&
            namespaceUri != null &&
            namespaceUri.contains('word')) {
          tableRow = current;
        } else if (current.name.local == 'tbl' &&
            namespaceUri != null &&
            namespaceUri.contains('word')) {
          table = current;
          break;
        }
        final parent = current.parent;
        current = parent is XmlElement ? parent : null;
      }

      if (tableCell == null || tableRow == null || table == null) {
        return {'row': null, 'column': null};
      }

      // Calculate row index: count preceding tr elements in the table
      int rowIndex = 0;
      for (var sibling in table.children) {
        if (sibling is XmlElement && sibling.name.local == 'tr') {
          if (sibling == tableRow) {
            break;
          }
          rowIndex++;
        }
      }

      // Calculate column index: count preceding tc elements in the row
      int columnIndex = 0;
      for (var sibling in tableRow.children) {
        if (sibling is XmlElement && sibling.name.local == 'tc') {
          if (sibling == tableCell) {
            break;
          }
          columnIndex++;
        }
      }

      return {'row': rowIndex, 'column': columnIndex};
    }

    // Helper function to build path
    String _buildPath(View view) {
      final pathSegments = <String>[];
      View? current = view;

      // Determine document section (document, header, footer)
      String section = 'document';
      final sdtElement = view.sdtView?.sdt;
      if (sdtElement != null) {
        XmlElement? xmlCurrent = sdtElement;
        while (xmlCurrent != null) {
          final parent = xmlCurrent.parent;
          if (parent is XmlElement) {
            final parentName = parent.name.toString();
            if (parentName.contains('header')) {
              section = 'header';
              break;
            } else if (parentName.contains('footer')) {
              section = 'footer';
              break;
            }
            xmlCurrent = parent;
          } else {
            break;
          }
        }
      }

      pathSegments.add(section);

      // Build path by traversing parent chain
      final viewChain = <View>[];
      while (current != null) {
        viewChain.insert(0, current);
        current = current.parentView;
      }

      // Skip root view
      for (var i = 1; i < viewChain.length; i++) {
        final v = viewChain[i];
        if (v is RowView) {
          // Find table index
          int tableIndex = 0;
          for (var j = i - 1; j >= 0; j--) {
            final prev = viewChain[j];
            if (prev is RowView && prev != v) {
              tableIndex++;
            } else if (prev is! RowView) {
              break;
            }
          }
          pathSegments.add('table[$tableIndex]');
        } else {
          // For other views, use a generic path segment
          final viewType =
              v.runtimeType.toString().replaceAll('View', '').toLowerCase();
          pathSegments.add('$viewType[${i - 1}]');
        }
      }

      // Add cell information if nested in table
      final parentRowView = _findParentRowView(view);
      if (parentRowView != null) {
        final indices = _calculateTableIndices(view);
        if (indices['row'] != null && indices['column'] != null) {
          pathSegments.add('row[${indices['row']}]');
          pathSegments.add('cell[${indices['column']}]');
        }
      }

      return pathSegments.join('/');
    }

    // Traverse all views
    void _collectTagsFromSub(Map<String, List<View>>? sub) {
      if (sub == null) return;

      for (var entry in sub.entries) {
        for (var view in entry.value) {
          final tagType = _getTagTypeFromView(view);
          final parentRowView = _findParentRowView(view);
          final isNested = parentRowView != null;
          final parentTableTag = parentRowView?.tag;
          final path = _buildPath(view);

          Map<String, int?> indices = {};
          int? rowIndex;
          int? columnIndex;

          if (isNested) {
            indices = _calculateTableIndices(view);
            rowIndex = indices['row'];
            columnIndex = indices['column'];
          }

          final docxTag = DocxTag(
            name: entry.key,
            type: tagType,
            isNested: isNested,
            parentTableTag: parentTableTag,
            columnIndex: columnIndex,
            rowIndex: rowIndex,
            path: path,
          );

          allTags.add(docxTag);

          // Group by type
          tagsByType.putIfAbsent(tagType, () => []).add(docxTag);

          // Group by nesting
          if (isNested && parentTableTag != null) {
            tableTags.putIfAbsent(parentTableTag, () => []).add(docxTag);
          } else {
            documentTags.add(docxTag);
          }

          // Recursively collect from nested views
          if (view.sub != null) {
            _collectTagsFromSub(view.sub);
          }
        }
      }
    }

    _collectTagsFromSub(viewManager.root.sub);

    return DocxTagCollection(
      allTags: allTags,
      tagsByType: tagsByType,
      tableTags: tableTags,
      documentTags: documentTags,
    );
  }

  ///
  ///Get all tags from template
  ///
  /// @Deprecated Use getTagsEnhanced() for better tag information including types and nesting
  @Deprecated('Use getTagsEnhanced() for better tag information')
  List<String> getTags() {
    return getTagsEnhanced().tagNames;
  }

  ///
  /// Generates byte buffer with docx file content by given [c]
  ///
  Future<List<int>?> generate(Content c,
      {TagPolicy tagPolicy = TagPolicy.saveText,
      ImagePolicy imagePolicy = ImagePolicy.save}) async {
    final vm = ViewManager.attach(_manager,
        tagPolicy: tagPolicy, imgPolicy: imagePolicy);
    vm.produce(c);
    _manager.updateArch();
    final enc = ZipEncoder();

    return enc.encode(_manager.arch);
  }

  // ---------------------------------------------------------------------------
  // Editing API
  //
  // The methods below operate on the DOCX as a static document — they do not
  // perform template filling. They let a caller take an arbitrary user-uploaded
  // DOCX, inspect its top-level body structure by stable indices, and rewrite
  // paragraphs and tables in place to insert SDT-based template tags. The
  // intended workflow is:
  //
  //   1. fromBytes(...) — load the user's DOCX
  //   2. getEditableStructure() — get indexed paragraphs & tables
  //   3. Make decisions about which content maps to which template tag
  //      (typically with help from an LLM)
  //   4. replaceParagraphText(...) / rewriteTableRows(...) — apply edits
  //   5. save() — get the rewritten DOCX bytes
  //
  // The output is intended to be consumed by the existing fill pipeline
  // (generate(...)), so the edits write proper SDT structures with alias/tag/id.
  // ---------------------------------------------------------------------------

  XmlElement get _body {
    final docEntry = _manager.getEntry(() => DocxXmlEntry(), 'word/document.xml');
    final doc = docEntry?.doc;
    if (doc == null) {
      throw DocxTemplateException('word/document.xml missing or unreadable');
    }
    final body = doc.rootElement.children
        .whereType<XmlElement>()
        .firstWhereOrNull((e) => e.name.local == 'body');
    if (body == null) {
      throw DocxTemplateException('word/document.xml has no <w:body>');
    }
    return body;
  }

  List<XmlElement> get _topLevelChildren =>
      _body.children.whereType<XmlElement>().toList();

  /// Snapshot the top-level body of the document so a caller can address
  /// paragraphs and tables by stable 0-based indices.
  DocxEditableStructure getEditableStructure() {
    final paragraphs = <DocxParagraphInfo>[];
    final tables = <DocxTableInfo>[];

    var pIdx = 0;
    var tIdx = 0;
    for (final el in _topLevelChildren) {
      switch (el.name.local) {
        case 'p':
          paragraphs.add(
            DocxParagraphInfo(pIdx: pIdx, text: _paragraphText(el)),
          );
          pIdx++;
          break;
        case 'tbl':
          tables.add(DocxTableInfo(tIdx: tIdx, rows: _tableRowText(el)));
          tIdx++;
          break;
        default:
          // Ignore sectPr and other body-level structural elements.
          break;
      }
    }

    return DocxEditableStructure(paragraphs: paragraphs, tables: tables);
  }

  /// Replace the text content of the paragraph at the given top-level index.
  ///
  /// Existing runs are removed; the paragraph's `<w:pPr>` is preserved. If
  /// [sdtTag] is provided, the new text is wrapped in a `<w:sdt>` block so the
  /// fill pipeline picks it up as a template tag; otherwise the text is written
  /// as plain runs.
  ///
  /// Throws if [pIdx] is out of range.
  void replaceParagraphText({
    required int pIdx,
    required String text,
    String? sdtTag,
    String? sdtAlias,
    SdtIdAllocator? idAllocator,
    String? backfillColorHex,
  }) {
    final paragraph = _findParagraph(pIdx);
    final pPr = paragraph.children
        .whereType<XmlElement>()
        .firstWhereOrNull((e) => e.name.local == 'pPr');

    // Inherit the first existing run's <w:rPr> so the new placeholder
    // run keeps the document's font/size — without this, Word renders
    // the SDT in its default font (typically Calibri 11), which stands
    // out against the body text. Optionally overlay a colour for
    // visual distinction (e.g. blue for backfill fields).
    final inheritedRpr = findFirstRunRpr(paragraph);
    final placeholderRpr = rPrWithColor(
      base: inheritedRpr,
      hexColour: backfillColorHex,
    );

    final newChildren = <XmlNode>[];
    if (pPr != null) newChildren.add(pPr.copy());

    if (sdtTag != null) {
      final allocator = idAllocator ?? SdtIdAllocator();
      newChildren.add(
        buildSdt(
          tag: sdtTag,
          alias: sdtAlias ?? sdtTag,
          id: allocator.next(),
          contentChildren: [buildRun(text: text, rPr: placeholderRpr)],
        ),
      );
    } else {
      newChildren.add(buildRun(text: text, rPr: placeholderRpr));
    }

    paragraph.children
      ..clear()
      ..addAll(newChildren);
  }

  /// Keep [keepHeaderRows] rows of the table at [tIdx], then replace a
  /// contiguous run of "data" rows starting at [keepHeaderRows] with a single
  /// SDT-wrapped templated row. Any rows beyond that data run are preserved
  /// verbatim, so this method is safe to use on tables that only partially
  /// expand at fill time (e.g. an inspection form where one big table holds
  /// header info, an irregularities sub-section, and trailing sign-off rows).
  ///
  /// The data run is determined by:
  ///   - [dataRows] when explicitly provided (takes precedence), OR
  ///   - auto-detection: walk forward from [keepHeaderRows] and stop at the
  ///     first row whose `<w:tc>` count differs from the first data row's
  ///     cell count. In typical Word forms, a section break is visible as a
  ///     row whose merged-cell count drops to 1 (or jumps), so this
  ///     heuristic correctly bounds the expandable section.
  ///
  /// The number of cells in the templated row should match the table's column
  /// count; if it doesn't, the row is written as-is and the table grid will
  /// determine final layout.
  ///
  /// Throws if [tIdx] is out of range or if the table has fewer than
  /// [keepHeaderRows] rows.
  void rewriteTableRows({
    required int tIdx,
    required int keepHeaderRows,
    required TemplatedRow templateRow,
    SdtIdAllocator? idAllocator,
    int? dataRows,
    String? backfillColorHex,
  }) {
    final table = _findTable(tIdx);
    final rows = table.children
        .whereType<XmlElement>()
        .where((e) => e.name.local == 'tr')
        .toList();
    if (rows.length < keepHeaderRows) {
      throw DocxTemplateException(
        'Table $tIdx has ${rows.length} rows, cannot keep $keepHeaderRows headers',
      );
    }

    // Use the first data row (or header) as the cell-shape template so we
    // preserve column widths via the existing <w:tcPr>/<w:trPr> properties.
    final shapeRow =
        rows.length > keepHeaderRows ? rows[keepHeaderRows] : rows.last;

    // Determine how many data rows belong to the expandable section.
    final detectedDataRows = _detectDataRowCount(
      rows: rows,
      keepHeaderRows: keepHeaderRows,
      shapeRow: shapeRow,
    );
    final effectiveDataRows = dataRows != null
        ? dataRows.clamp(0, rows.length - keepHeaderRows)
        : detectedDataRows;
    final dataEnd = keepHeaderRows + effectiveDataRows;

    final allocator = idAllocator ?? SdtIdAllocator();
    final newDataRow = _buildTemplatedRow(
      shapeRow: shapeRow,
      templateRow: templateRow,
      idAllocator: allocator,
      backfillColorHex: backfillColorHex,
    );

    // The wrapper SDT must use the literal tag value "table" so the fill
    // pipeline (ViewManager._processSdt) classifies it as a RowView and
    // expands the templated row once per content item at fill time.
    // The data binding (e.g. "step/1/checkitems") goes in the alias.
    // Setting tag to anything else makes it a TextView, which renders only
    // a single inert row and silently drops your repeating data.
    final wrappedRow = buildSdt(
      tag: 'table',
      alias: templateRow.wrapperAlias,
      id: allocator.next(),
      contentChildren: [newDataRow],
    );

    // Walk children in order. Replace tr rows in [keepHeaderRows, dataEnd)
    // with a single wrapped row at the start of that range; keep header rows
    // and any rows beyond dataEnd untouched. Non-tr children (tblPr, tblGrid,
    // etc.) are preserved in place.
    final out = <XmlNode>[];
    var trCount = 0;
    var inserted = false;
    for (final node in table.children) {
      if (node is XmlElement && node.name.local == 'tr') {
        if (trCount < keepHeaderRows) {
          out.add(node.copy());
        } else if (trCount < dataEnd) {
          if (!inserted) {
            out.add(wrappedRow);
            inserted = true;
          }
          // Otherwise drop — this row is replaced by the wrapper.
        } else {
          // Trailing row beyond the data section — keep as-is.
          out.add(node.copy());
        }
        trCount++;
      } else {
        out.add(node.copy());
      }
    }
    // Edge case: empty data section (effectiveDataRows == 0) — append the
    // wrapper after any existing rows so the binding still resolves.
    if (!inserted) out.add(wrappedRow);

    table.children
      ..clear()
      ..addAll(out);
  }

  /// Count the contiguous run of rows starting at [keepHeaderRows] whose
  /// `<w:tc>` count matches the [shapeRow]'s cell count. Returns at least 1
  /// when there is room for a data row, so the templated wrapper always has
  /// somewhere to live.
  int _detectDataRowCount({
    required List<XmlElement> rows,
    required int keepHeaderRows,
    required XmlElement shapeRow,
  }) {
    if (keepHeaderRows >= rows.length) return 0;
    final shapeCellCount = shapeRow.children
        .whereType<XmlElement>()
        .where((e) => e.name.local == 'tc')
        .length;
    var count = 0;
    for (var i = keepHeaderRows; i < rows.length; i++) {
      final cellCount = rows[i]
          .children
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'tc')
          .length;
      if (cellCount != shapeCellCount) break;
      count++;
    }
    return count == 0 ? 1 : count;
  }

  /// Replace the content of one cell in [tIdx]/[rowIdx] at [cellIdx]. The
  /// cell's `<w:tcPr>` (column properties) is preserved; existing paragraphs
  /// inside the cell are dropped and a single new paragraph is written. If
  /// [sdtTag] is provided the new paragraph contains an SDT (so the existing
  /// fill pipeline picks it up); otherwise it is plain text.
  ///
  /// Use this for inspection-style tables where each step lives as one row
  /// of a shared table and only specific empty cells (e.g. the Pass/Fail
  /// column) should receive backfill placeholders.
  ///
  /// Throws if any index is out of range.
  void replaceCellContent({
    required int tIdx,
    required int rowIdx,
    required int cellIdx,
    required String text,
    String? sdtTag,
    String? sdtAlias,
    SdtIdAllocator? idAllocator,
    String? backfillColorHex,
  }) {
    final table = _findTable(tIdx);
    final rows = table.children
        .whereType<XmlElement>()
        .where((e) => e.name.local == 'tr')
        .toList();
    if (rowIdx < 0 || rowIdx >= rows.length) {
      throw DocxTemplateException(
        'Row index $rowIdx out of range (table $tIdx has ${rows.length} rows)',
      );
    }
    final row = rows[rowIdx];
    final cells = row.children
        .whereType<XmlElement>()
        .where((e) => e.name.local == 'tc')
        .toList();
    if (cellIdx < 0 || cellIdx >= cells.length) {
      throw DocxTemplateException(
        'Cell index $cellIdx out of range (row $rowIdx has ${cells.length} cells)',
      );
    }
    final cell = cells[cellIdx];

    // Preserve <w:tcPr> if present; everything else (paragraphs, nested
    // tables) is replaced with a single new paragraph carrying the placeholder.
    final tcPr = cell.children
        .whereType<XmlElement>()
        .firstWhereOrNull((e) => e.name.local == 'tcPr');

    // Inherit run properties from the cell's first existing run so the
    // placeholder text keeps the cell's font/size, then optionally
    // overlay the backfill colour.
    final inheritedRpr = findFirstRunRpr(cell);
    final placeholderRpr = rPrWithColor(
      base: inheritedRpr,
      hexColour: backfillColorHex,
    );

    final w = (String local) => XmlName(local, 'w');
    final XmlElement paragraphContent;
    if (sdtTag != null) {
      final allocator = idAllocator ?? SdtIdAllocator();
      paragraphContent = XmlElement(w('p'), [], [
        buildSdt(
          tag: sdtTag,
          alias: sdtAlias ?? sdtTag,
          id: allocator.next(),
          contentChildren: [buildRun(text: text, rPr: placeholderRpr)],
        ),
      ]);
    } else {
      paragraphContent = XmlElement(w('p'), [], [
        buildRun(text: text, rPr: placeholderRpr),
      ]);
    }

    cell.children
      ..clear()
      ..addAll([
        if (tcPr != null) tcPr.copy(),
        paragraphContent,
      ]);
  }

  /// Like [replaceCellContent] but writes an *image* SDT instead of a text
  /// one — for cells that should resolve to an image at fill time
  /// (e.g. signee/signature). Adds a placeholder PNG to the archive and a
  /// matching relationship under `word/_rels/document.xml.rels`, then
  /// inserts an SDT containing a properly-shaped `<w:drawing>` element so
  /// the fill pipeline classifies it as an `ImgView`.
  ///
  /// The placeholder image is replaced by the real bytes at fill time;
  /// the size in EMU here defines the rendered box size.
  /// (914400 EMU = 1 inch.) Defaults give a ~2 in × 0.5 in signature box.
  ///
  /// Throws if any index is out of range or if `word/_rels/document.xml.rels`
  /// is missing.
  void replaceCellContentWithImage({
    required int tIdx,
    required int rowIdx,
    required int cellIdx,
    required String sdtAlias,
    String sdtTag = 'img',
    int widthEmu = 1828800,
    int heightEmu = 457200,
    SdtIdAllocator? idAllocator,
  }) {
    final relsEntry = _manager.getEntry(
      () => DocxRelsEntry(),
      'word/_rels/document.xml.rels',
    );
    if (relsEntry == null) {
      throw DocxTemplateException(
        'word/_rels/document.xml.rels missing — cannot add image relationship',
      );
    }

    final imageId = relsEntry.nextImageId();
    final relId = relsEntry.nextId();
    final mediaPath = 'word/media/$imageId.png';
    relsEntry.add(
      relId,
      DocxRel(
        relId,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/$imageId.png',
      ),
    );
    _manager.add(mediaPath, DocxBinEntry(kPlaceholderPngBytes));

    final allocator = idAllocator ?? SdtIdAllocator();
    final imgSdt = buildImageSdt(
      tag: sdtTag,
      alias: sdtAlias,
      id: allocator.next(),
      relId: relId,
      widthEmu: widthEmu,
      heightEmu: heightEmu,
      docPrId: int.parse(imageId.replaceAll('image', '')),
    );

    // Replace the cell's children using the same logic as
    // replaceCellContent, but with the image SDT. Preserve <w:tcPr>.
    final table = _findTable(tIdx);
    final rows = table.children
        .whereType<XmlElement>()
        .where((e) => e.name.local == 'tr')
        .toList();
    if (rowIdx < 0 || rowIdx >= rows.length) {
      throw DocxTemplateException(
        'Row index $rowIdx out of range (table $tIdx has ${rows.length} rows)',
      );
    }
    final row = rows[rowIdx];
    final cells = row.children
        .whereType<XmlElement>()
        .where((e) => e.name.local == 'tc')
        .toList();
    if (cellIdx < 0 || cellIdx >= cells.length) {
      throw DocxTemplateException(
        'Cell index $cellIdx out of range (row $rowIdx has ${cells.length} cells)',
      );
    }
    final cell = cells[cellIdx];
    final tcPr = cell.children
        .whereType<XmlElement>()
        .firstWhereOrNull((e) => e.name.local == 'tcPr');

    cell.children
      ..clear()
      ..addAll([
        if (tcPr != null) tcPr.copy(),
        imgSdt,
      ]);
  }

  /// Re-encode the (potentially edited) archive into DOCX bytes. Use this
  /// instead of [generate] when you have only edited the document, not run
  /// template filling.
  Future<List<int>?> save() async {
    _manager.updateArch();
    return ZipEncoder().encode(_manager.arch);
  }

  XmlElement _findParagraph(int pIdx) {
    var i = 0;
    for (final el in _topLevelChildren) {
      if (el.name.local == 'p') {
        if (i == pIdx) return el;
        i++;
      }
    }
    throw DocxTemplateException('Paragraph index $pIdx out of range ($i)');
  }

  XmlElement _findTable(int tIdx) {
    var i = 0;
    for (final el in _topLevelChildren) {
      if (el.name.local == 'tbl') {
        if (i == tIdx) return el;
        i++;
      }
    }
    throw DocxTemplateException('Table index $tIdx out of range ($i)');
  }

  String _paragraphText(XmlElement paragraph) {
    final buf = StringBuffer();
    for (final t in paragraph.descendants.whereType<XmlElement>()) {
      if (t.name.local == 't') buf.write(t.innerText);
    }
    return buf.toString();
  }

  List<List<String>> _tableRowText(XmlElement table) {
    final rows = <List<String>>[];
    for (final tr in table.children.whereType<XmlElement>()) {
      if (tr.name.local != 'tr') continue;
      final cells = <String>[];
      for (final tc in tr.children.whereType<XmlElement>()) {
        if (tc.name.local != 'tc') continue;
        final buf = StringBuffer();
        for (final t in tc.descendants.whereType<XmlElement>()) {
          if (t.name.local == 't') {
            if (buf.isNotEmpty) buf.write(' ');
            buf.write(t.innerText);
          }
        }
        cells.add(buf.toString());
      }
      rows.add(cells);
    }
    return rows;
  }

  XmlElement _buildTemplatedRow({
    required XmlElement shapeRow,
    required TemplatedRow templateRow,
    required SdtIdAllocator idAllocator,
    String? backfillColorHex,
  }) {
    final w = (String local) => XmlName(local, 'w');

    // Copy <w:trPr> if present so row height/header settings are preserved.
    final trPr = shapeRow.children
        .whereType<XmlElement>()
        .firstWhereOrNull((e) => e.name.local == 'trPr');

    final shapeCells = shapeRow.children
        .whereType<XmlElement>()
        .where((e) => e.name.local == 'tc')
        .toList();

    final newCells = <XmlElement>[];
    for (var i = 0; i < templateRow.cells.length; i++) {
      final cellRecipe = templateRow.cells[i];
      // Preserve <w:tcPr> from the matching shape cell so column widths stay.
      final shapeCell = i < shapeCells.length ? shapeCells[i] : null;
      final shapeTcPr = shapeCell != null
          ? shapeCell.children
              .whereType<XmlElement>()
              .firstWhereOrNull((e) => e.name.local == 'tcPr')
          : null;

      // Inherit run properties from the shape cell so per-column font
      // and size choices carry through into the per-iteration runs the
      // fill pipeline emits. Optionally overlay the backfill colour.
      final inheritedRpr =
          shapeCell != null ? findFirstRunRpr(shapeCell) : null;
      final placeholderRpr = rPrWithColor(
        base: inheritedRpr,
        hexColour: backfillColorHex,
      );

      final cellChildren = <XmlNode>[];
      if (shapeTcPr != null) cellChildren.add(shapeTcPr.copy());

      // Inner cells inside a `tag="table"` wrapper must themselves carry
      // a type marker in the SDT tag attribute so the fill pipeline can
      // classify them under the parent RowView. Use "text" for normal
      // cells (TextView resolves text bindings) and "img" for image
      // cells (ImgView swaps the placeholder for ImageContent bytes).
      // The cell's data binding (e.g. "col/1/text", "complete", "date",
      // "signee/signature") goes in the alias — that's what the fill
      // pipeline routes content against.
      final XmlElement paragraphContent;
      if (cellRecipe.image) {
        // Add a placeholder PNG to the archive and a fresh image rel so
        // the drawing has something concrete to point at. The fill
        // pipeline replaces both at fill time.
        final relsEntry = _manager.getEntry(
          () => DocxRelsEntry(),
          'word/_rels/document.xml.rels',
        );
        if (relsEntry == null) {
          throw DocxTemplateException(
            'word/_rels/document.xml.rels missing — cannot add image cell',
          );
        }
        final imageId = relsEntry.nextImageId();
        final relId = relsEntry.nextId();
        relsEntry.add(
          relId,
          DocxRel(
            relId,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'media/$imageId.png',
          ),
        );
        _manager.add(
          'word/media/$imageId.png',
          DocxBinEntry(kPlaceholderPngBytes),
        );
        paragraphContent = buildImageSdt(
          tag: 'img',
          alias: cellRecipe.tag,
          id: idAllocator.next(),
          relId: relId,
          widthEmu: cellRecipe.imageWidthEmu,
          heightEmu: cellRecipe.imageHeightEmu,
          docPrId: int.parse(imageId.replaceAll('image', '')),
        );
      } else {
        paragraphContent = buildSdt(
          tag: 'text',
          alias: cellRecipe.tag,
          id: idAllocator.next(),
          contentChildren: [
            buildRun(
              text: cellRecipe.placeholder ?? '',
              rPr: placeholderRpr,
            ),
          ],
        );
      }

      cellChildren.add(XmlElement(w('p'), [], [paragraphContent]));
      newCells.add(XmlElement(w('tc'), [], cellChildren));
    }

    return XmlElement(w('tr'), [], [
      if (trPr != null) trPr.copy(),
      ...newCells,
    ]);
  }
}

