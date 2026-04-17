import 'dart:io';

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

      // Wrapper SDT with tag step/1/checkitems wraps a single tr.
      final sdt = table.children
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'sdt');
      final wrapperTag = sdt.descendants
          .whereType<XmlElement>()
          .firstWhere((e) => e.name.local == 'tag');
      expect(wrapperTag.getAttribute('val', namespace: '*'),
          'step/1/checkitems');

      final innerRows = sdt.descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'tr')
          .toList();
      expect(innerRows.length, 1);

      // Two cells, each with an SDT for text & complete.
      final cellTags = sdt.descendants
          .whereType<XmlElement>()
          .where((e) => e.name.local == 'tag')
          .map((e) => e.getAttribute('val', namespace: '*'))
          .toList();
      expect(cellTags, containsAll(['step/1/checkitems', 'text', 'complete']));
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
