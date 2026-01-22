/// Represents the type of a tag in a DOCX template
enum TagType {
  /// Simple text replacement
  text,

  /// Table tag that will be replaced with table content
  table,

  /// Image tag
  image,

  /// Dropdown/combobox field
  combobox,

  /// Checkbox field
  checkbox,

  /// Date field
  date,

  /// Rich text with formatting
  richText,

  /// List tag
  list,

  /// Plain block tag
  plain,

  /// Hyperlink tag
  link,
}

/// Represents a tag found in a DOCX template
class DocxTag {
  /// The tag name (e.g., "projectTitle", "step/1/title")
  final String name;

  /// The type of tag
  final TagType type;

  /// Whether this tag is nested inside a table
  final bool isNested;

  /// If nested, the parent table tag name (if applicable)
  final String? parentTableTag;

  /// If nested, the column index within the table (0-based)
  final int? columnIndex;

  /// If nested, the row index within the table (0-based, -1 for header row)
  final int? rowIndex;

  /// Full path to the tag in the document structure (e.g., "document/body/table[0]/row[1]/cell[2]")
  final String path;

  DocxTag({
    required this.name,
    required this.type,
    this.isNested = false,
    this.parentTableTag,
    this.columnIndex,
    this.rowIndex,
    required this.path,
  });

  @override
  String toString() => name;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is DocxTag &&
          runtimeType == other.runtimeType &&
          name == other.name &&
          type == other.type &&
          path == other.path;

  @override
  int get hashCode => name.hashCode ^ type.hashCode ^ path.hashCode;
}

/// Collection of tags with metadata
class DocxTagCollection {
  /// All tags found in the document
  final List<DocxTag> allTags;

  /// Tags grouped by type
  final Map<TagType, List<DocxTag>> tagsByType;

  /// Table tags and their nested tags (key: table tag name, value: nested tags)
  final Map<String, List<DocxTag>> tableTags;

  /// Document-level tags (not nested in tables)
  final List<DocxTag> documentTags;

  DocxTagCollection({
    required this.allTags,
    required this.tagsByType,
    required this.tableTags,
    required this.documentTags,
  });

  /// Get all tag names (for backward compatibility)
  List<String> get tagNames => allTags.map((t) => t.name).toList();

  /// Get tags of a specific type
  List<DocxTag> getTagsByType(TagType type) => tagsByType[type] ?? [];

  /// Get nested tags for a specific table
  List<DocxTag> getNestedTagsForTable(String tableTagName) =>
      tableTags[tableTagName] ?? [];
}
