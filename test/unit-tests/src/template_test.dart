import 'dart:io';

import 'package:docx_template/docx_template.dart';
import 'package:test/test.dart';

void main() {
  test('getTags', () async {
    final f = File("template.docx");
    final docx = await DocxTemplate.fromBytes(await f.readAsBytes());
    final list = docx.getTags();
    // print(list);
    expect(list.length, 12);
    expect(list.first, 'imgFirst');
    expect(list[1], 'docname');
    expect(list[2], 'list');
    expect(list[3], 'table');
    expect(list[4], 'passport');
    expect(list[5], 'plainlist');
    expect(list[6], 'multilineList');
    expect(list[7], 'multilineText2');
    expect(list[8], 'img');
    expect(list[9], 'link');
    expect(list[10], 'header');
    expect(list[11], 'logo');
  });

  test('getTagsEnhanced - basic functionality', () async {
    final f = File("template.docx");
    final docx = await DocxTemplate.fromBytes(await f.readAsBytes());
    final tagCollection = docx.getTagsEnhanced();

    // Should have the same number of tags as getTags()
    expect(tagCollection.allTags.length, 12);
    expect(tagCollection.tagNames.length, 12);

    // Verify all tag names are present
    final tagNames = tagCollection.tagNames.toSet();
    expect(tagNames.contains('imgFirst'), isTrue);
    expect(tagNames.contains('docname'), isTrue);
    expect(tagNames.contains('list'), isTrue);
    expect(tagNames.contains('table'), isTrue);
  });

  test('getTagsEnhanced - tag type detection', () async {
    final f = File("template.docx");
    final docx = await DocxTemplate.fromBytes(await f.readAsBytes());
    final tagCollection = docx.getTagsEnhanced();

    // Find specific tags and verify their types
    final imgFirstTag =
        tagCollection.allTags.firstWhere((t) => t.name == 'imgFirst');
    expect(imgFirstTag.type, TagType.image);

    final docnameTag =
        tagCollection.allTags.firstWhere((t) => t.name == 'docname');
    expect(docnameTag.type, TagType.text);

    final listTag = tagCollection.allTags.firstWhere((t) => t.name == 'list');
    expect(listTag.type, TagType.list);

    final tableTag = tagCollection.allTags.firstWhere((t) => t.name == 'table');
    expect(tableTag.type, TagType.table);

    final linkTag = tagCollection.allTags.firstWhere((t) => t.name == 'link');
    expect(linkTag.type, TagType.link);
  });

  test('getTagsEnhanced - tags by type', () async {
    final f = File("template.docx");
    final docx = await DocxTemplate.fromBytes(await f.readAsBytes());
    final tagCollection = docx.getTagsEnhanced();

    // Get tags by type
    final imageTags = tagCollection.getTagsByType(TagType.image);
    expect(imageTags.length, greaterThan(0));
    expect(imageTags.any((t) => t.name == 'imgFirst'), isTrue);

    final textTags = tagCollection.getTagsByType(TagType.text);
    expect(textTags.length, greaterThan(0));
    expect(textTags.any((t) => t.name == 'docname'), isTrue);

    final tableTags = tagCollection.getTagsByType(TagType.table);
    expect(tableTags.length, greaterThan(0));
    expect(tableTags.any((t) => t.name == 'table'), isTrue);
  });

  test('getTagsEnhanced - nested tag detection', () async {
    final f = File("template.docx");
    final docx = await DocxTemplate.fromBytes(await f.readAsBytes());
    final tagCollection = docx.getTagsEnhanced();

    // Get nested tags for the table
    final nestedTags = tagCollection.getNestedTagsForTable('table');

    // If there are nested tags, verify they have correct properties
    if (nestedTags.isNotEmpty) {
      for (var tag in nestedTags) {
        expect(tag.isNested, isTrue);
        expect(tag.parentTableTag, 'table');
        expect(tag.columnIndex, isNotNull);
        expect(tag.rowIndex, isNotNull);
      }
    }

    // Verify tableTags map contains the table
    expect(tagCollection.tableTags.containsKey('table'), isTrue);
  });

  test('getTagsEnhanced - document tags vs nested tags', () async {
    final f = File("template.docx");
    final docx = await DocxTemplate.fromBytes(await f.readAsBytes());
    final tagCollection = docx.getTagsEnhanced();

    // Document tags should not be nested
    for (var tag in tagCollection.documentTags) {
      expect(tag.isNested, isFalse);
      expect(tag.parentTableTag, isNull);
    }

    // All tags should be either document tags or nested in tables
    final totalDocumentAndNested = tagCollection.documentTags.length +
        tagCollection.tableTags.values
            .fold(0, (sum, tags) => sum + tags.length);
    expect(totalDocumentAndNested,
        greaterThanOrEqualTo(tagCollection.allTags.length));
  });

  test('getTagsEnhanced - path building', () async {
    final f = File("template.docx");
    final docx = await DocxTemplate.fromBytes(await f.readAsBytes());
    final tagCollection = docx.getTagsEnhanced();

    // All tags should have a path
    for (var tag in tagCollection.allTags) {
      expect(tag.path, isNotEmpty);
      expect(tag.path.contains('/'), isTrue);
    }

    // Document tags should have paths starting with document, header, or footer
    for (var tag in tagCollection.documentTags) {
      expect(
        tag.path.startsWith('document') ||
            tag.path.startsWith('header') ||
            tag.path.startsWith('footer'),
        isTrue,
      );
    }
  });

  test('getTagsEnhanced - backward compatibility', () async {
    final f = File("template.docx");
    final docx = await DocxTemplate.fromBytes(await f.readAsBytes());

    // getTags() should return the same as getTagsEnhanced().tagNames
    final oldTags = docx.getTags();
    final newTags = docx.getTagsEnhanced().tagNames;

    expect(oldTags.length, newTags.length);
    expect(oldTags.toSet(), equals(newTags.toSet()));
  });

  // test('generate pdf', () async {
  //   final f = File("template.docx");
  //   final docx = await DocxTemplate.fromBytes(await f.readAsBytes());
  //   final list = docx.exportPdf();
  // });
}
