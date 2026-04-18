import 'package:archive/archive.dart';
import 'package:docx_template/docx_template.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

/// Build a minimal valid DOCX in-memory with the given <w:body> children
/// (XML strings). Returns the bytes for use with DocxTemplate.fromBytes.
List<int> _buildDocx(String bodyChildren) {
  const wNs = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  final document =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<w:document $wNs><w:body>$bodyChildren</w:body></w:document>';

  const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>';

  const rootRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      '</Relationships>';

  final arch = Archive();
  void put(String name, String content) {
    final bytes = content.codeUnits;
    arch.addFile(ArchiveFile(name, bytes.length, bytes));
  }

  put('[Content_Types].xml', contentTypes);
  put('_rels/.rels', rootRels);
  put('word/document.xml', document);

  return ZipEncoder().encode(arch);
}

XmlDocument _readDocumentXml(List<int> bytes) {
  final arch = ZipDecoder().decodeBytes(bytes);
  final f = arch.files.firstWhere((f) => f.name == 'word/document.xml');
  return XmlDocument.parse(String.fromCharCodes(f.content as List<int>));
}

void main() {
  group('DocxTemplate editing API', () {
    test('getEditableStructure indexes top-level paragraphs and tables', () async {
      final bytes = _buildDocx(
        '<w:p><w:r><w:t>Title</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>Para two</w:t></w:r></w:p>'
        '<w:tbl>'
        '<w:tr><w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>'
        '</w:tbl>'
        '<w:p><w:r><w:t>After table</w:t></w:r></w:p>',
      );

      final docx = await DocxTemplate.fromBytes(bytes);
      final structure = docx.getEditableStructure();

      expect(structure.paragraphs.length, 3);
      expect(structure.paragraphs[0].pIdx, 0);
      expect(structure.paragraphs[0].text, 'Title');
      expect(structure.paragraphs[1].text, 'Para two');
      expect(structure.paragraphs[2].text, 'After table');

      expect(structure.tables.length, 1);
      expect(structure.tables[0].tIdx, 0);
      expect(structure.tables[0].rows, [
        ['H1', 'H2'],
        ['A', 'B'],
      ]);
    });

    test('replaceParagraphText replaces runs and adds SDT when tagged', () async {
      final bytes = _buildDocx(
        '<w:p><w:r><w:t>Original</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>Other</w:t></w:r></w:p>',
      );

      final docx = await DocxTemplate.fromBytes(bytes);
      docx.replaceParagraphText(pIdx: 0, text: 'New plain');
      docx.replaceParagraphText(pIdx: 1, text: 'placeholder', sdtTag: 'title');

      final out = await docx.save();
      expect(out, isNotNull);
      final xml = _readDocumentXml(out!);

      final body = xml
          .descendants
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'body');
      final paragraphs = body.children
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'p')
          .toList();
      expect(paragraphs.length, 2);

      // Plain replacement.
      final p0Text = paragraphs[0]
          .descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 't')
          .map((e) => e.innerText)
          .join();
      expect(p0Text, 'New plain');

      // Tagged paragraph: should contain a w:sdt with w:tag val=title.
      final sdts = paragraphs[1]
          .descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'sdt')
          .toList();
      expect(sdts, isNotEmpty);
      final tagEl = sdts.first.descendants
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'tag');
      expect(tagEl.getAttribute('val', namespace: '*'), 'title');
    });

    test('rewriteTableRows keeps header and inserts wrapped templated row',
        () async {
      final bytes = _buildDocx(
        '<w:tbl>'
        '<w:tr><w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>Old A</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>Old B</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>Old C</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>Old D</w:t></w:r></w:p></w:tc></w:tr>'
        '</w:tbl>',
      );

      final docx = await DocxTemplate.fromBytes(bytes);
      docx.rewriteTableRows(
        tIdx: 0,
        keepHeaderRows: 1,
        templateRow: TemplatedRow(
          wrapperTag: 'step/1/checkitems',
          cells: [
            TemplatedCell(tag: 'text'),
            TemplatedCell(tag: 'complete'),
          ],
        ),
      );

      final out = await docx.save();
      expect(out, isNotNull);
      final xml = _readDocumentXml(out!);

      final table = xml
          .descendants
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'tbl');

      // Direct <w:tr> children: header only (other data rows replaced).
      final directRows = table.children
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'tr')
          .toList();
      expect(directRows.length, 1, reason: 'old data rows must be removed');

      // Wrapper SDT must use tag="table" (RowView marker) and put the
      // data binding (step/1/checkitems) in the alias. Without the
      // tag="table" the fill pipeline classifies it as a TextView and
      // the templated row never repeats per check item.
      final sdt = table.children
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'sdt');
      final wrapperSdtPr = sdt.children
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'sdtPr');
      final wrapperTag = wrapperSdtPr.children
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'tag');
      final wrapperAlias = wrapperSdtPr.children
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'alias');
      expect(wrapperTag.getAttribute('val', namespace: '*'), 'table');
      expect(wrapperAlias.getAttribute('val', namespace: '*'),
          'step/1/checkitems');

      final innerRows = sdt.descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'tr')
          .toList();
      expect(innerRows.length, 1);

      // Inner cell SDTs use tag="text" with the data binding in alias.
      final cellAliases = innerRows.first.descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'alias')
          .map((e) => e.getAttribute('val', namespace: '*'))
          .toList();
      expect(cellAliases, containsAll(['text', 'complete']));
    });

    test('round-trips: edited DOCX is recognised by getTagsEnhanced', () async {
      final bytes = _buildDocx(
        '<w:p><w:r><w:t>Will become title</w:t></w:r></w:p>',
      );

      final docx = await DocxTemplate.fromBytes(bytes);
      docx.replaceParagraphText(pIdx: 0, text: '', sdtTag: 'title');
      final out = await docx.save();

      final reloaded = await DocxTemplate.fromBytes(out!);
      final tags = reloaded.getTagsEnhanced();
      expect(tags.allTags.map((t) => t.name).toList(), contains('title'));
    });

    test('replaceCellContent tags one cell, keeps siblings intact', () async {
      final bytes = _buildDocx(
        '<w:tbl>'
        '<w:tr><w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>Item one</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>Item two</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc></w:tr>'
        '</w:tbl>',
      );

      final docx = await DocxTemplate.fromBytes(bytes);
      docx.replaceCellContent(
        tIdx: 0,
        rowIdx: 1,
        cellIdx: 1,
        text: '',
        sdtTag: 'step/1/complete',
      );
      docx.replaceCellContent(
        tIdx: 0,
        rowIdx: 2,
        cellIdx: 1,
        text: '',
        sdtTag: 'step/2/complete',
      );

      final out = await docx.save();
      final reloaded = await DocxTemplate.fromBytes(out!);
      final tagNames =
          reloaded.getTagsEnhanced().allTags.map((t) => t.name).toList();
      expect(tagNames, contains('step/1/complete'));
      expect(tagNames, contains('step/2/complete'));

      // Description column (cell 0) untouched — still contains literal text.
      final xml = _readDocumentXml(out);
      final body = xml.descendants
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'body');
      final dataRows = body
          .descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'tr')
          .skip(1)
          .toList();
      expect(dataRows.length, 2);
      final firstDescCell = dataRows[0]
          .children
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'tc')
          .first;
      final descText = firstDescCell
          .descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 't')
          .map((e) => e.innerText)
          .join();
      expect(descText, 'Item one');
    });

    test('fill works on edited template even when placeholder text was empty',
        () async {
      // Regression: writing replaceParagraphText/replaceCellContent with
      // text='' would produce <w:t xml:space="preserve"/> which becomes a
      // self-closing tag with NO child text node after save → reload. The
      // fill pipeline previously did `t.children[0] = XmlText(value)` and
      // crashed with RangeError. _setTText now handles this case.
      final bytes = _buildDocx(
        '<w:p><w:r><w:t>Will become title</w:t></w:r></w:p>',
      );
      final docx = await DocxTemplate.fromBytes(bytes);
      docx.replaceParagraphText(pIdx: 0, text: '', sdtTag: 'title');
      final templateBytes = (await docx.save())!;

      // Round-trip: load the edited template and fill it as if at shift time.
      final filled = await DocxTemplate.fromBytes(templateBytes);
      final content = Content()..add(TextContent('title', 'Lubricate Bearings'));
      final outBytes = await filled.generate(content);
      expect(outBytes, isNotNull);

      // Verify the fill substituted the value.
      final reloaded = await DocxTemplate.fromBytes(outBytes!);
      final allText = ZipDecoder()
          .decodeBytes(outBytes)
          .files
          .firstWhere((f) => f.name == 'word/document.xml');
      final xml = String.fromCharCodes(allText.content as List<int>);
      expect(xml, contains('Lubricate Bearings'));
      // Suppress unused-variable for the reload integrity check.
      expect(reloaded.getTagsEnhanced().allTags, isNotEmpty);
    });

    test('post-fill cleanup strips bare <text>/<table> View leftovers',
        () async {
      // Regression: ViewManager replaces SDTs with View elements whose
      // XmlName has no namespace prefix (e.g. "text"). When some Views
      // never get fully replaced by produce() (nested-under-TextView,
      // missing content keys, etc.) they used to leak into the output as
      // bare `<text>` elements — which Word rejects as malformed OOXML.
      // _stripLeakedViewElements scrubs them at the end of produce.
      //
      // We trigger the leak by injecting a bare <text> element directly
      // and verifying generate() removes it.
      final bytes = _buildDocx(
        '<w:p><w:r><w:t>Before</w:t></w:r></w:p>'
        '<w:p><text><w:r><w:t>Inner</w:t></w:r></text></w:p>'
        '<w:p><w:r><w:t>After</w:t></w:r></w:p>',
      );
      final docx = await DocxTemplate.fromBytes(bytes);
      final out = await docx.generate(Content());
      final reloadedXml = String.fromCharCodes(
        ZipDecoder()
            .decodeBytes(out!)
            .files
            .firstWhere((f) => f.name == 'word/document.xml')
            .content as List<int>,
      );
      expect(
        reloadedXml,
        isNot(contains('<text>')),
        reason: 'bare <text> View leaks must be unwrapped before saving',
      );
      expect(reloadedXml, contains('Inner'));
    });

    test(
        'rewriteTableRows preserves trailing rows when shape changes '
        '(inspection-form pattern)', () async {
      // Mirror the ITF_12.docx pattern: a single table holds header info,
      // a 4-cell irregularities sub-section, and trailing 1-cell sign-off
      // rows. Only the contiguous 4-cell data run should be replaced; the
      // 1-cell trailing rows must survive into the output.
      final bytes = _buildDocx(
        '<w:tbl>'
        // Header row (4 cells) — keep.
        '<w:tr>'
        '<w:tc><w:p><w:r><w:t>Date</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>Time</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>Description</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>Signature</w:t></w:r></w:p></w:tc>'
        '</w:tr>'
        // Two 4-cell data rows — replace with single wrapper.
        '<w:tr>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>'
        '</w:tr>'
        '<w:tr>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>'
        '</w:tr>'
        // Trailing 1-cell rows (different shape) — keep verbatim.
        '<w:tr>'
        '<w:tc><w:p><w:r><w:t>Approved by Manager</w:t></w:r></w:p></w:tc>'
        '</w:tr>'
        '<w:tr>'
        '<w:tc><w:p><w:r><w:t>Signature line</w:t></w:r></w:p></w:tc>'
        '</w:tr>'
        '</w:tbl>',
      );

      final docx = await DocxTemplate.fromBytes(bytes);
      docx.rewriteTableRows(
        tIdx: 0,
        keepHeaderRows: 1,
        templateRow: TemplatedRow(
          wrapperTag: 'step/1/rows',
          cells: [
            TemplatedCell(tag: 'date'),
            TemplatedCell(tag: 'time'),
            TemplatedCell(tag: 'col/1/text'),
            TemplatedCell(tag: 'signee/name'),
          ],
        ),
      );

      final out = await docx.save();
      expect(out, isNotNull);
      final xml = _readDocumentXml(out!);
      final table = xml
          .descendants
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'tbl');

      // Direct rows: 1 header + 1 wrapper SDT (not a tr) + 2 trailing trs.
      // The wrapper SDT is not a w:tr child, so direct tr children = 3.
      final directRows = table.children
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'tr')
          .toList();
      expect(directRows.length, 3,
          reason:
              'header + 2 trailing single-cell rows should survive; the two '
              '4-cell data rows are collapsed into one wrapper SDT');

      // Wrapper SDT exists and uses tag="table" with the right alias.
      final sdt = table.children
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'sdt');
      final wrapperAlias = sdt.descendants
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'alias');
      expect(wrapperAlias.getAttribute('val', namespace: '*'), 'step/1/rows');

      // Trailing rows still carry their literal text — this is the
      // regression we're guarding against.
      final allText = table.descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 't')
          .map((e) => e.innerText)
          .join(' ');
      expect(allText, contains('Approved by Manager'));
      expect(allText, contains('Signature line'));
    });

    test('rewriteTableRows respects explicit dataRows override', () async {
      // When the AI knows the data run length precisely, dataRows takes
      // precedence over the auto-detected same-shape window.
      final bytes = _buildDocx(
        '<w:tbl>'
        '<w:tr><w:tc><w:p><w:r><w:t>H</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>D1</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>D2</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:tc></w:tr>'
        '</w:tbl>',
      );
      final docx = await DocxTemplate.fromBytes(bytes);
      docx.rewriteTableRows(
        tIdx: 0,
        keepHeaderRows: 1,
        dataRows: 2,
        templateRow: TemplatedRow(
          wrapperTag: 'step/1/rows',
          cells: [TemplatedCell(tag: 'text')],
        ),
      );
      final out = await docx.save();
      final xml = _readDocumentXml(out!);
      final allText = xml.descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 't')
          .map((e) => e.innerText)
          .join(' ');
      expect(allText, contains('H'));
      expect(allText, contains('Keep'));
      expect(allText, isNot(contains('D1')));
      expect(allText, isNot(contains('D2')));
    });

    test('out-of-range indices throw', () async {
      final bytes = _buildDocx(
        '<w:p><w:r><w:t>only</w:t></w:r></w:p>',
      );
      final docx = await DocxTemplate.fromBytes(bytes);
      expect(
        () => docx.replaceParagraphText(pIdx: 5, text: 'x'),
        throwsA(isA<DocxTemplateException>()),
      );
      expect(
        () => docx.rewriteTableRows(
          tIdx: 0,
          keepHeaderRows: 0,
          templateRow: TemplatedRow(
            wrapperTag: 'foo',
            cells: [TemplatedCell(tag: 'bar')],
          ),
        ),
        throwsA(isA<DocxTemplateException>()),
      );
    });
  });
}
