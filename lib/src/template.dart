import 'package:archive/archive.dart';
import 'package:collection/collection.dart' show IterableExtension;
import 'package:docx_template/docx_template.dart';
import 'package:docx_template/src/tag_models.dart';
import 'package:docx_template/src/view_manager.dart';
import 'package:xml/xml.dart';

import 'docx_entry.dart';

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
}
