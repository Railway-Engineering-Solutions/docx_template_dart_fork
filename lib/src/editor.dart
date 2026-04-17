import 'package:xml/xml.dart';

/// Snapshot of the top-level body structure of a DOCX, addressable by stable
/// 0-based indices for rewriting.
///
/// Only top-level body elements are indexed: paragraphs (`<w:p>`) and tables
/// (`<w:tbl>`) appearing directly under `<w:body>`. Nested content (a paragraph
/// inside a table cell) is not in [paragraphs] — table cells are addressable
/// via [tables] indices.
class DocxEditableStructure {
  DocxEditableStructure({
    required this.paragraphs,
    required this.tables,
  });

  final List<DocxParagraphInfo> paragraphs;
  final List<DocxTableInfo> tables;
}

class DocxParagraphInfo {
  DocxParagraphInfo({required this.pIdx, required this.text});

  /// 0-based index of this paragraph among top-level body paragraphs in
  /// document order.
  final int pIdx;

  /// Concatenated text content of all `<w:t>` runs in the paragraph.
  final String text;
}

class DocxTableInfo {
  DocxTableInfo({required this.tIdx, required this.rows});

  /// 0-based index of this table among top-level body tables in document order.
  final int tIdx;

  /// Cell text per row: `rows[r][c]` is the concatenated text of cell `c` in
  /// row `r`.
  final List<List<String>> rows;
}

/// One cell in a [TemplatedRow]. The cell will be written as a paragraph
/// containing a single SDT with the given tag/alias and placeholder text.
class TemplatedCell {
  /// If [alias] is omitted it defaults to [tag]. The existing fill pipeline
  /// keys lookups off the alias, so for tags containing `/` the alias must
  /// match the full path verbatim or the tag won't resolve at fill time.
  const TemplatedCell({
    required this.tag,
    this.alias,
    this.placeholder,
  });

  final String tag;
  final String? alias;
  final String? placeholder;

  String get effectiveAlias => alias ?? tag;
}

/// Recipe for the single data row inserted by [DocxTemplate.rewriteTableRows].
///
/// The row is wrapped in an outer SDT (`wrapperTag` / `wrapperAlias`) so that
/// the existing template engine recognises it as a table tag and repeats the
/// row per content item at fill time. If [wrapperAlias] is omitted it
/// defaults to [wrapperTag] — the fill pipeline keys lookups off the alias,
/// so they must match for paths containing `/`.
class TemplatedRow {
  TemplatedRow({
    required this.wrapperTag,
    String? wrapperAlias,
    required this.cells,
  }) : wrapperAlias = wrapperAlias ?? wrapperTag;

  final String wrapperTag;
  final String wrapperAlias;
  final List<TemplatedCell> cells;
}

/// Allocates monotonically increasing SDT id values within an editing session.
/// Negative ids match the convention used by Word for generated SDTs.
class SdtIdAllocator {
  int _next = -1000000;

  String next() {
    final v = _next.toString();
    _next++;
    return v;
  }
}

/// Builds a `<w:sdt>` element wrapping [contentChildren] with the given tag
/// metadata. The `w` namespace prefix is expected to be bound at the document
/// root (it always is in valid DOCX files).
XmlElement buildSdt({
  required String tag,
  required String alias,
  required String id,
  required List<XmlNode> contentChildren,
}) {
  XmlName w(String local) => XmlName(local, 'w');

  XmlElement valEl(String local, String val) => XmlElement(
        w(local),
        [XmlAttribute(w('val'), val)],
      );

  final sdtPr = XmlElement(w('sdtPr'), [], [
    valEl('alias', alias),
    valEl('tag', tag),
    valEl('id', id),
  ]);

  final sdtContent = XmlElement(w('sdtContent'), [], contentChildren);

  return XmlElement(w('sdt'), [], [sdtPr, sdtContent]);
}

/// Builds a `<w:r><w:t xml:space="preserve">text</w:t></w:r>` run.
XmlElement buildRun({required String text, XmlElement? rPr}) {
  XmlName w(String local) => XmlName(local, 'w');
  final t = XmlElement(
    w('t'),
    [XmlAttribute(XmlName('space', 'xml'), 'preserve')],
    [XmlText(text)],
  );
  return XmlElement(w('r'), [], [
    if (rPr != null) rPr,
    t,
  ]);
}

/// Builds a `<w:p>` containing a single run with [text]. Used when generating
/// fresh paragraphs (e.g. for templated table cells).
XmlElement buildParagraph({required String text}) {
  XmlName w(String local) => XmlName(local, 'w');
  return XmlElement(w('p'), [], [buildRun(text: text)]);
}
